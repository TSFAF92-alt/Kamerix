Attribute VB_Name = "ZmodFormBuilder"
Option Explicit
Option Private Module

'==============================================================================
' KAMERIX - Outil : construction des ecrans de connexion (userforms)
'------------------------------------------------------------------------------
' BuildKamerixForms (F5) cree, ou recree s'ils existent deja :
'   frmLogin          : connexion (avec le lien "Mot de passe oublie ?")
'   frmFirstAdmin     : premier lancement, creation de l'administrateur
'   frmChangePassword : changement du mot de passe (obligatoire ou volontaire)
'   frmRecovery       : code de secours, redefinition du mot de passe d'un administrateur
'   frmUsers          : gestion des comptes
'   frmUserEdit       : nouveau compte ou changement de role
'   frmReference      : referentiel (marques, modeles, villes, quartiers, commissariats)
'   frmParties        : tiers et intermediaires (liste) ; frmPartyEdit, frmIntermediaryEdit (fiches)
'   frmCars           : voitures (liste) ; frmCarEdit (nouvel achat ou fiche voiture)
'   frmSales          : ventes comptant (liste) ; frmSaleStart, frmSalePayment, frmSaleCv, frmSaleCancel
'   frmSaleDocuments  : documents d'une vente (facture, recus, remboursements), edition PDF ou impression
'   frmSettings       : parametres a onglets (entreprise, vente, securite) ; assistant de premiere configuration
' Chaque formulaire existant est d'abord exporte dans Backup_VBA_aaaammjj_hhnnss,
' puis remplace. Le code des formulaires est ecrit par ce module : pour le
' modifier, changer ce module puis relancer BuildKamerixForms.
' Apres usage : Debogage > Compiler VBAProject.
' Prerequis   : "Acces approuve au modele d'objet du projet VBA" coche.
' Liaison     : tardive (aucune reference VBIDE necessaire).
'==============================================================================

Private Const VBEXT_MS_FORM As Long = 3
Private Const VBEXT_PP_LOCKED As Long = 1
Private Const MAX_MODULE_NAME As Long = 31
Private Const TOOL_TITLE As String = "Kamerix - construction des ecrans"

Private Const FONT_NAME As String = "Segoe UI"
Private Const FONT_SIZE As Single = 9
Private Const MARGIN As Single = 12
Private Const LABEL_WIDTH As Single = 96
Private Const FIELD_WIDTH As Single = 168
Private Const ROW_HEIGHT As Single = 26
Private Const BUTTON_WIDTH As Single = 84
Private Const BUTTON_HEIGHT As Single = 24
Private Const FORM_EXTRA_HEIGHT As Single = 28     ' barre de titre et bordures

Private mstrBackupFolder As String

'==============================================================================
' API PUBLIQUE
'==============================================================================

Public Sub BuildKamerixForms()
    Dim objProject As Object
    Dim strReport As String

    On Error GoTo EH

    Set objProject = GetProject()
    If objProject Is Nothing Then Exit Sub
    mstrBackupFolder = ""

    BuildLoginForm objProject
    strReport = strReport & "OK     frmLogin" & vbCrLf
    BuildFirstAdminForm objProject
    strReport = strReport & "OK     frmFirstAdmin" & vbCrLf
    BuildChangePasswordForm objProject
    strReport = strReport & "OK     frmChangePassword" & vbCrLf
    BuildRecoveryForm objProject
    strReport = strReport & "OK     frmRecovery" & vbCrLf
    BuildUsersForm objProject
    strReport = strReport & "OK     frmUsers" & vbCrLf
    BuildUserEditForm objProject
    strReport = strReport & "OK     frmUserEdit" & vbCrLf
    BuildReferenceForm objProject
    strReport = strReport & "OK     frmReference" & vbCrLf
    BuildPartiesForm objProject
    strReport = strReport & "OK     frmParties" & vbCrLf
    BuildPartyEditForm objProject
    strReport = strReport & "OK     frmPartyEdit" & vbCrLf
    BuildIntermediaryEditForm objProject
    strReport = strReport & "OK     frmIntermediaryEdit" & vbCrLf
    BuildCarsForm objProject
    strReport = strReport & "OK     frmCars" & vbCrLf
    BuildCarEditForm objProject
    strReport = strReport & "OK     frmCarEdit" & vbCrLf
    BuildSalesForm objProject
    strReport = strReport & "OK     frmSales" & vbCrLf
    BuildSaleStartForm objProject
    strReport = strReport & "OK     frmSaleStart" & vbCrLf
    BuildSalePaymentForm objProject
    strReport = strReport & "OK     frmSalePayment" & vbCrLf
    BuildSaleCvForm objProject
    strReport = strReport & "OK     frmSaleCv" & vbCrLf
    BuildSaleCancelForm objProject
    strReport = strReport & "OK     frmSaleCancel" & vbCrLf
    BuildSaleDocumentsForm objProject
    strReport = strReport & "OK     frmSaleDocuments" & vbCrLf
    BuildSettingsForm objProject
    strReport = strReport & "OK     frmSettings" & vbCrLf

    If Len(mstrBackupFolder) > 0 Then
        strReport = strReport & vbCrLf & "Anciens formulaires sauvegardes dans :" & vbCrLf & mstrBackupFolder & vbCrLf
    End If
    strReport = strReport & vbCrLf & "Etape suivante : Debogage > Compiler VBAProject, puis lancer StartKamerix."

    Debug.Print strReport
    MsgBox strReport, vbInformation, TOOL_TITLE
    Exit Sub

EH:
    MsgBox "Echec de la construction (erreur " & Err.Number & ") :" & vbCrLf & Err.Description & _
           vbCrLf & vbCrLf & "Deja construits :" & vbCrLf & strReport, vbCritical, TOOL_TITLE
End Sub

'==============================================================================
' MISE EN PAGE DES FORMULAIRES
'==============================================================================

Private Sub BuildLoginForm(ByVal objProject As Object)
    Dim objForm As Object
    Dim sngTop As Single
    Dim sngWidth As Single

    sngWidth = MARGIN * 2 + LABEL_WIDTH + FIELD_WIDTH
    Set objForm = NewForm(objProject, "frmLogin", "Kamerix - Connexion")

    AddTitle objForm, "lblTitle", "Connexion " & ChrW(224) & " Kamerix", sngWidth, MARGIN
    sngTop = MARGIN + 30
    AddField objForm, "lblLogin", "Identifiant :", "txtLogin", False, sngTop
    sngTop = sngTop + ROW_HEIGHT
    AddField objForm, "lblPassword", "Mot de passe :", "txtPassword", True, sngTop
    sngTop = sngTop + ROW_HEIGHT
    AddLink objForm, "lblForgot", "Mot de passe oubli" & ChrW(233) & " ?", MARGIN + LABEL_WIDTH, sngTop - 2, FIELD_WIDTH
    sngTop = sngTop + 18
    AddMessage objForm, sngWidth, sngTop
    sngTop = sngTop + 34
    AddButtons objForm, "Se connecter", "Annuler", sngWidth, sngTop

    SizeForm objForm, sngWidth, sngTop + BUTTON_HEIGHT + MARGIN
    SetFormCode objForm, CodeFrmLogin()
End Sub

Private Sub BuildFirstAdminForm(ByVal objProject As Object)
    Dim objForm As Object
    Dim sngTop As Single
    Dim sngWidth As Single

    sngWidth = MARGIN * 2 + LABEL_WIDTH + FIELD_WIDTH
    Set objForm = NewForm(objProject, "frmFirstAdmin", "Kamerix - Premier lancement")

    AddTitle objForm, "lblTitle", "Cr" & ChrW(233) & "ation de l'administrateur", sngWidth, MARGIN
    AddInfo objForm, "lblInfo", "Aucun compte n'existe encore. Cr" & ChrW(233) & "ez le compte administrateur. " & _
        "Son mot de passe ne pourra pas " & ChrW(234) & "tre relu : notez-le en lieu s" & ChrW(251) & "r.", _
        sngWidth, MARGIN + 28, 30
    sngTop = MARGIN + 66
    AddField objForm, "lblLogin", "Identifiant :", "txtLogin", False, sngTop
    sngTop = sngTop + ROW_HEIGHT
    AddField objForm, "lblFullName", "Nom complet :", "txtFullName", False, sngTop
    sngTop = sngTop + ROW_HEIGHT
    AddField objForm, "lblPassword", "Mot de passe :", "txtPassword", True, sngTop
    sngTop = sngTop + ROW_HEIGHT
    AddField objForm, "lblConfirm", "Confirmation :", "txtConfirm", True, sngTop
    sngTop = sngTop + ROW_HEIGHT
    AddInfo objForm, "lblRules", "", sngWidth, sngTop, 16
    sngTop = sngTop + 20
    AddMessage objForm, sngWidth, sngTop
    sngTop = sngTop + 34
    AddButtons objForm, "Cr" & ChrW(233) & "er", "Annuler", sngWidth, sngTop

    SizeForm objForm, sngWidth, sngTop + BUTTON_HEIGHT + MARGIN
    SetFormCode objForm, CodeFrmFirstAdmin()
End Sub

Private Sub BuildChangePasswordForm(ByVal objProject As Object)
    Dim objForm As Object
    Dim sngTop As Single
    Dim sngWidth As Single

    sngWidth = MARGIN * 2 + LABEL_WIDTH + FIELD_WIDTH
    Set objForm = NewForm(objProject, "frmChangePassword", "Kamerix - Mot de passe")

    AddTitle objForm, "lblTitle", "Changement du mot de passe", sngWidth, MARGIN
    AddInfo objForm, "lblInfo", "", sngWidth, MARGIN + 28, 30
    sngTop = MARGIN + 66
    AddField objForm, "lblOld", "Mot de passe actuel :", "txtOld", True, sngTop
    sngTop = sngTop + ROW_HEIGHT
    AddField objForm, "lblNew", "Nouveau :", "txtNew", True, sngTop
    sngTop = sngTop + ROW_HEIGHT
    AddField objForm, "lblConfirm", "Confirmation :", "txtConfirm", True, sngTop
    sngTop = sngTop + ROW_HEIGHT
    AddInfo objForm, "lblRules", "", sngWidth, sngTop, 16
    sngTop = sngTop + 20
    AddMessage objForm, sngWidth, sngTop
    sngTop = sngTop + 34
    AddButtons objForm, "Valider", "Annuler", sngWidth, sngTop

    SizeForm objForm, sngWidth, sngTop + BUTTON_HEIGHT + MARGIN
    SetFormCode objForm, CodeFrmChangePassword()
End Sub

Private Sub BuildRecoveryForm(ByVal objProject As Object)
    Dim objForm As Object
    Dim sngTop As Single
    Dim sngWidth As Single

    sngWidth = MARGIN * 2 + LABEL_WIDTH + FIELD_WIDTH
    Set objForm = NewForm(objProject, "frmRecovery", "Kamerix - Code de secours")

    AddTitle objForm, "lblTitle", "Mot de passe oubli" & ChrW(233), sngWidth, MARGIN
    AddInfo objForm, "lblInfo", "R" & ChrW(233) & "serv" & ChrW(233) & " aux administrateurs : saisissez le code de secours " & _
        "remis " & ChrW(224) & " la cr" & ChrW(233) & "ation du premier administrateur.", sngWidth, MARGIN + 28, 30
    sngTop = MARGIN + 66
    AddField objForm, "lblCode", "Code de secours :", "txtCode", False, sngTop
    sngTop = sngTop + ROW_HEIGHT
    AddField objForm, "lblLogin", "Identifiant :", "txtLogin", False, sngTop
    sngTop = sngTop + ROW_HEIGHT
    AddField objForm, "lblNew", "Nouveau :", "txtNew", True, sngTop
    sngTop = sngTop + ROW_HEIGHT
    AddField objForm, "lblConfirm", "Confirmation :", "txtConfirm", True, sngTop
    sngTop = sngTop + ROW_HEIGHT
    AddInfo objForm, "lblRules", "", sngWidth, sngTop, 16
    sngTop = sngTop + 20
    AddMessage objForm, sngWidth, sngTop
    sngTop = sngTop + 34
    AddButtons objForm, "Valider", "Annuler", sngWidth, sngTop

    SizeForm objForm, sngWidth, sngTop + BUTTON_HEIGHT + MARGIN
    SetFormCode objForm, CodeFrmRecovery()
End Sub

Private Sub BuildUsersForm(ByVal objProject As Object)
    Const LIST_WIDTH As Single = 400
    Const LIST_HEIGHT As Single = 216
    Const SIDE_WIDTH As Single = 150
    Dim objForm As Object
    Dim ctl As Object
    Dim sngWidth As Single
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim vNames As Variant
    Dim vCaptions As Variant
    Dim i As Long

    sngWidth = MARGIN * 3 + LIST_WIDTH + SIDE_WIDTH
    Set objForm = NewForm(objProject, "frmUsers", "Kamerix - Comptes")

    AddTitle objForm, "lblTitle", "Gestion des comptes", sngWidth, MARGIN

    ' En-tetes des colonnes (memes largeurs que lstUsers.ColumnWidths : 80, 130, 85, 90)
    sngTop = MARGIN + 30
    AddLabelAt objForm, "lblColLogin", "Identifiant", MARGIN + 3, sngTop, 80
    AddLabelAt objForm, "lblColName", "Nom", MARGIN + 83, sngTop, 130
    AddLabelAt objForm, "lblColRole", "R" & ChrW(244) & "le", MARGIN + 213, sngTop, 85
    AddLabelAt objForm, "lblColState", ChrW(201) & "tat", MARGIN + 298, sngTop, 90

    sngTop = sngTop + 16
    Set ctl = AddControl(objForm, "Forms.ListBox.1", "lstUsers", MARGIN, sngTop, LIST_WIDTH, LIST_HEIGHT)

    ' Boutons d'action, a droite de la liste
    sngLeft = MARGIN * 2 + LIST_WIDTH
    vNames = Array("cmdNew", "cmdRole", "cmdReset", "cmdUnlock", "cmdActive", "cmdRecovery")
    vCaptions = Array("Nouveau compte...", "Changer le r" & ChrW(244) & "le...", "Nouveau mot de passe", _
                      "D" & ChrW(233) & "bloquer", "D" & ChrW(233) & "sactiver", "Nouveau code de secours")
    For i = LBound(vNames) To UBound(vNames)
        AddButtonAt objForm, CStr(vNames(i)), CStr(vCaptions(i)), sngLeft, sngTop + i * (BUTTON_HEIGHT + 6), SIDE_WIDTH
    Next i

    sngTop = sngTop + LIST_HEIGHT + 6
    AddMessage objForm, sngWidth, sngTop
    sngTop = sngTop + 34

    Set ctl = AddControl(objForm, "Forms.CommandButton.1", "cmdClose", sngWidth - MARGIN - BUTTON_WIDTH, sngTop, _
                         BUTTON_WIDTH, BUTTON_HEIGHT)
    ctl.Caption = "Fermer"
    ctl.Cancel = True

    SizeForm objForm, sngWidth, sngTop + BUTTON_HEIGHT + MARGIN
    SetFormCode objForm, CodeFrmUsers()
End Sub

Private Sub BuildUserEditForm(ByVal objProject As Object)
    Dim objForm As Object
    Dim ctl As Object
    Dim sngTop As Single
    Dim sngWidth As Single

    sngWidth = MARGIN * 2 + LABEL_WIDTH + FIELD_WIDTH
    Set objForm = NewForm(objProject, "frmUserEdit", "Kamerix - Compte")

    AddTitle objForm, "lblTitle", "Compte", sngWidth, MARGIN
    AddInfo objForm, "lblInfo", "", sngWidth, MARGIN + 28, 30
    sngTop = MARGIN + 66
    AddField objForm, "lblLogin", "Identifiant :", "txtLogin", False, sngTop
    sngTop = sngTop + ROW_HEIGHT
    AddField objForm, "lblFullName", "Nom complet :", "txtFullName", False, sngTop
    sngTop = sngTop + ROW_HEIGHT
    AddLabelAt objForm, "lblRole", "R" & ChrW(244) & "le :", MARGIN, sngTop + 3, LABEL_WIDTH
    Set ctl = AddControl(objForm, "Forms.ComboBox.1", "cboRole", MARGIN + LABEL_WIDTH, sngTop, FIELD_WIDTH, 20)
    sngTop = sngTop + ROW_HEIGHT
    AddMessage objForm, sngWidth, sngTop
    sngTop = sngTop + 34
    AddButtons objForm, "Valider", "Annuler", sngWidth, sngTop

    SizeForm objForm, sngWidth, sngTop + BUTTON_HEIGHT + MARGIN
    SetFormCode objForm, CodeFrmUserEdit()
End Sub

Private Sub BuildReferenceForm(ByVal objProject As Object)
    Const LIST_WIDTH As Single = 270
    Const LIST_HEIGHT As Single = 200
    Const SIDE_WIDTH As Single = 120
    Dim objForm As Object
    Dim ctl As Object
    Dim sngWidth As Single
    Dim sngTop As Single
    Dim sngLeft As Single

    sngWidth = MARGIN * 3 + LIST_WIDTH + SIDE_WIDTH
    Set objForm = NewForm(objProject, "frmReference", "Kamerix - R" & ChrW(233) & "f" & ChrW(233) & "rentiel")

    AddTitle objForm, "lblTitle", "R" & ChrW(233) & "f" & ChrW(233) & "rentiel", sngWidth, MARGIN
    sngTop = MARGIN + 30
    AddLabelAt objForm, "lblKind", "Liste :", MARGIN, sngTop + 3, 50
    Set ctl = AddControl(objForm, "Forms.ComboBox.1", "cboKind", MARGIN + 50, sngTop, 130, 20)
    AddLabelAt objForm, "lblParent", "Marque :", MARGIN + 195, sngTop + 3, 50
    Set ctl = AddControl(objForm, "Forms.ComboBox.1", "cboParent", MARGIN + 245, sngTop, sngWidth - MARGIN * 2 - 245, 20)

    ' En-tetes des colonnes (memes largeurs que lstItems.ColumnWidths : 160 puis 90, ou 250)
    sngTop = sngTop + 30
    AddLabelAt objForm, "lblCol1", "Marque", MARGIN + 3, sngTop, 160
    AddLabelAt objForm, "lblCol2", "Ville", MARGIN + 163, sngTop, 90
    sngTop = sngTop + 16
    Set ctl = AddControl(objForm, "Forms.ListBox.1", "lstItems", MARGIN, sngTop, LIST_WIDTH, LIST_HEIGHT)

    sngLeft = MARGIN * 2 + LIST_WIDTH
    AddLabelAt objForm, "lblName", "Nom :", sngLeft, sngTop, SIDE_WIDTH
    Set ctl = AddControl(objForm, "Forms.TextBox.1", "txtName", sngLeft, sngTop + 16, SIDE_WIDTH, 20)
    AddButtonAt objForm, "cmdAdd", "Ajouter", sngLeft, sngTop + 48, SIDE_WIDTH
    AddButtonAt objForm, "cmdRename", "Renommer", sngLeft, sngTop + 48 + (BUTTON_HEIGHT + 6), SIDE_WIDTH
    AddButtonAt objForm, "cmdDelete", "Supprimer", sngLeft, sngTop + 48 + (BUTTON_HEIGHT + 6) * 2, SIDE_WIDTH

    sngTop = sngTop + LIST_HEIGHT + 6
    AddMessage objForm, sngWidth, sngTop
    sngTop = sngTop + 34
    Set ctl = AddControl(objForm, "Forms.CommandButton.1", "cmdClose", sngWidth - MARGIN - BUTTON_WIDTH, sngTop, _
                         BUTTON_WIDTH, BUTTON_HEIGHT)
    ctl.Caption = "Fermer"
    ctl.Cancel = True

    SizeForm objForm, sngWidth, sngTop + BUTTON_HEIGHT + MARGIN
    SetFormCode objForm, CodeFrmReference()
End Sub

Private Sub BuildPartiesForm(ByVal objProject As Object)
    Const LIST_WIDTH As Single = 480
    Const LIST_HEIGHT As Single = 220
    Const SIDE_WIDTH As Single = 110
    Dim objForm As Object
    Dim ctl As Object
    Dim sngWidth As Single
    Dim sngTop As Single
    Dim sngLeft As Single

    sngWidth = MARGIN * 3 + LIST_WIDTH + SIDE_WIDTH
    Set objForm = NewForm(objProject, "frmParties", "Kamerix - Tiers")
    AddTitle objForm, "lblTitle", "Tiers et interm" & ChrW(233) & "diaires", sngWidth, MARGIN

    sngTop = MARGIN + 30
    AddLabelAt objForm, "lblKind", "Liste :", MARGIN, sngTop + 3, 40
    Set ctl = AddControl(objForm, "Forms.ComboBox.1", "cboKind", MARGIN + 40, sngTop, 140, 20)
    AddLabelAt objForm, "lblSearch", "Rechercher :", MARGIN + 195, sngTop + 3, 70
    Set ctl = AddControl(objForm, "Forms.TextBox.1", "txtSearch", MARGIN + 265, sngTop, 120, 20)
    AddLabelAt objForm, "lblRole", "R" & ChrW(244) & "le :", MARGIN + 400, sngTop + 3, 35
    Set ctl = AddControl(objForm, "Forms.ComboBox.1", "cboRole", MARGIN + 435, sngTop, 110, 20)

    ' En-tetes (largeurs de lstItems : 160, 95, 110, 105)
    sngTop = sngTop + 30
    AddLabelAt objForm, "lblCol1", "Nom", MARGIN + 3, sngTop, 160
    AddLabelAt objForm, "lblCol2", "T" & ChrW(233) & "l" & ChrW(233) & "phone", MARGIN + 163, sngTop, 95
    AddLabelAt objForm, "lblCol3", "Pi" & ChrW(232) & "ce", MARGIN + 258, sngTop, 110
    AddLabelAt objForm, "lblCol4", "R" & ChrW(244) & "le", MARGIN + 368, sngTop, 105
    sngTop = sngTop + 16
    Set ctl = AddControl(objForm, "Forms.ListBox.1", "lstItems", MARGIN, sngTop, LIST_WIDTH, LIST_HEIGHT)

    sngLeft = MARGIN * 2 + LIST_WIDTH
    AddButtonAt objForm, "cmdNew", "Nouveau...", sngLeft, sngTop, SIDE_WIDTH
    AddButtonAt objForm, "cmdEdit", "Modifier...", sngLeft, sngTop + BUTTON_HEIGHT + 6, SIDE_WIDTH

    sngTop = sngTop + LIST_HEIGHT + 6
    AddMessage objForm, sngWidth, sngTop
    sngTop = sngTop + 34
    AddCloseButton objForm, sngWidth, sngTop
    SizeForm objForm, sngWidth, sngTop + BUTTON_HEIGHT + MARGIN
    SetFormCode objForm, CodeFrmParties()
End Sub

Private Sub BuildPartyEditForm(ByVal objProject As Object)
    Dim objForm As Object
    Dim ctl As Object
    Dim sngTop As Single
    Dim sngWidth As Single

    sngWidth = MARGIN * 2 + LABEL_WIDTH + FIELD_WIDTH
    Set objForm = NewForm(objProject, "frmPartyEdit", "Kamerix - Tiers")
    AddTitle objForm, "lblTitle", "Tiers", sngWidth, MARGIN
    sngTop = MARGIN + 34
    AddField objForm, "lblName", "Nom :", "txtName", False, sngTop
    sngTop = sngTop + ROW_HEIGHT
    AddField objForm, "lblPhone", "T" & ChrW(233) & "l" & ChrW(233) & "phone :", "txtPhone", False, sngTop
    sngTop = sngTop + ROW_HEIGHT
    AddField objForm, "lblIdType", "Type de pi" & ChrW(232) & "ce :", "txtIdType", False, sngTop
    sngTop = sngTop + ROW_HEIGHT
    AddField objForm, "lblIdNumber", "N" & ChrW(176) & " de pi" & ChrW(232) & "ce :", "txtIdNumber", False, sngTop
    sngTop = sngTop + ROW_HEIGHT
    AddField objForm, "lblAddress", "Adresse :", "txtAddress", False, sngTop
    sngTop = sngTop + ROW_HEIGHT
    AddLabelAt objForm, "lblRoles", "R" & ChrW(244) & "le :", MARGIN, sngTop + 3, LABEL_WIDTH
    AddCheckBox objForm, "chkBuyer", "Acheteur", MARGIN + LABEL_WIDTH, sngTop, 80
    AddCheckBox objForm, "chkSeller", "Vendeur", MARGIN + LABEL_WIDTH + 85, sngTop, 80
    sngTop = sngTop + ROW_HEIGHT
    AddMessage objForm, sngWidth, sngTop
    sngTop = sngTop + 34
    AddButtons objForm, "Enregistrer", "Annuler", sngWidth, sngTop
    SizeForm objForm, sngWidth, sngTop + BUTTON_HEIGHT + MARGIN
    SetFormCode objForm, CodeFrmPartyEdit()
End Sub

Private Sub BuildIntermediaryEditForm(ByVal objProject As Object)
    Dim objForm As Object
    Dim ctl As Object
    Dim sngTop As Single
    Dim sngWidth As Single

    sngWidth = MARGIN * 2 + LABEL_WIDTH + FIELD_WIDTH
    Set objForm = NewForm(objProject, "frmIntermediaryEdit", "Kamerix - Interm" & ChrW(233) & "diaire")
    AddTitle objForm, "lblTitle", "Interm" & ChrW(233) & "diaire", sngWidth, MARGIN
    sngTop = MARGIN + 34
    AddField objForm, "lblName", "Nom :", "txtName", False, sngTop
    sngTop = sngTop + ROW_HEIGHT
    AddField objForm, "lblPhone", "T" & ChrW(233) & "l" & ChrW(233) & "phone :", "txtPhone", False, sngTop
    sngTop = sngTop + ROW_HEIGHT
    AddLabelAt objForm, "lblCity", "Ville :", MARGIN, sngTop + 3, LABEL_WIDTH
    Set ctl = AddControl(objForm, "Forms.ComboBox.1", "cboCity", MARGIN + LABEL_WIDTH, sngTop, FIELD_WIDTH, 20)
    sngTop = sngTop + ROW_HEIGHT
    AddMessage objForm, sngWidth, sngTop
    sngTop = sngTop + 34
    AddButtons objForm, "Enregistrer", "Annuler", sngWidth, sngTop
    SizeForm objForm, sngWidth, sngTop + BUTTON_HEIGHT + MARGIN
    SetFormCode objForm, CodeFrmIntermediaryEdit()
End Sub

Private Sub BuildCarsForm(ByVal objProject As Object)
    Const LIST_WIDTH As Single = 520
    Const LIST_HEIGHT As Single = 230
    Const SIDE_WIDTH As Single = 120
    Dim objForm As Object
    Dim ctl As Object
    Dim sngWidth As Single
    Dim sngTop As Single
    Dim sngLeft As Single

    sngWidth = MARGIN * 3 + LIST_WIDTH + SIDE_WIDTH
    Set objForm = NewForm(objProject, "frmCars", "Kamerix - Voitures")
    AddTitle objForm, "lblTitle", "Voitures", sngWidth, MARGIN

    sngTop = MARGIN + 30
    AddLabelAt objForm, "lblStatus", "Statut :", MARGIN, sngTop + 3, 45
    Set ctl = AddControl(objForm, "Forms.ComboBox.1", "cboStatus", MARGIN + 45, sngTop, 130, 20)
    AddLabelAt objForm, "lblCount", "", MARGIN + 190, sngTop + 3, 150

    ' En-tetes (largeurs de lstCars : 125, 35, 115, 85, 80, 70)
    sngTop = sngTop + 30
    AddLabelAt objForm, "lblCol1", "Mod" & ChrW(232) & "le", MARGIN + 3, sngTop, 125
    AddLabelAt objForm, "lblCol2", "Ann" & ChrW(233) & "e", MARGIN + 128, sngTop, 35
    AddLabelAt objForm, "lblCol3", "Ch" & ChrW(226) & "ssis", MARGIN + 163, sngTop, 115
    AddLabelAt objForm, "lblCol4", "Immatriculation", MARGIN + 278, sngTop, 85
    AddLabelAt objForm, "lblCol5", "Statut", MARGIN + 363, sngTop, 80
    AddLabelAt objForm, "lblCol6", "Prix d'achat", MARGIN + 443, sngTop, 70
    sngTop = sngTop + 16
    Set ctl = AddControl(objForm, "Forms.ListBox.1", "lstCars", MARGIN, sngTop, LIST_WIDTH, LIST_HEIGHT)

    sngLeft = MARGIN * 2 + LIST_WIDTH
    AddButtonAt objForm, "cmdPurchase", "Nouvel achat...", sngLeft, sngTop, SIDE_WIDTH
    AddButtonAt objForm, "cmdEdit", "Modifier...", sngLeft, sngTop + BUTTON_HEIGHT + 6, SIDE_WIDTH
    AddButtonAt objForm, "cmdWithdraw", "Retirer du stock...", sngLeft, sngTop + (BUTTON_HEIGHT + 6) * 2, SIDE_WIDTH

    sngTop = sngTop + LIST_HEIGHT + 6
    AddMessage objForm, sngWidth, sngTop
    sngTop = sngTop + 34
    AddCloseButton objForm, sngWidth, sngTop
    SizeForm objForm, sngWidth, sngTop + BUTTON_HEIGHT + MARGIN
    SetFormCode objForm, CodeFrmCars()
End Sub

Private Sub BuildCarEditForm(ByVal objProject As Object)
    Const COL_LABEL As Single = 85
    Const COL_FIELD As Single = 160
    Dim objForm As Object
    Dim ctl As Object
    Dim sngWidth As Single
    Dim sngTop As Single
    Dim sngRight As Single
    Dim i As Long
    Dim vLeft As Variant
    Dim vRight As Variant

    sngRight = MARGIN * 2 + COL_LABEL + COL_FIELD + 12
    sngWidth = sngRight + COL_LABEL + COL_FIELD + MARGIN
    Set objForm = NewForm(objProject, "frmCarEdit", "Kamerix - Voiture")
    AddTitle objForm, "lblTitle", "Voiture", sngWidth, MARGIN

    sngTop = MARGIN + 30
    AddSection objForm, "lblSectionCar", "Voiture", MARGIN, sngTop, COL_LABEL + COL_FIELD
    AddSection objForm, "lblSectionPurchase", "Achat", sngRight, sngTop, COL_LABEL + COL_FIELD
    sngTop = sngTop + 22

    ' Colonne de gauche : la voiture ; colonne de droite : l'achat. Chaque element :
    ' nom du libelle | texte | nom du controle | type (C = liste deroulante, T = texte)
    vLeft = Array("lblBrand|Marque :|cboBrand|C", "lblModel|Mod" & ChrW(232) & "le :|cboModel|C", _
                  "lblYear|Ann" & ChrW(233) & "e :|txtYear|T", "lblColor|Couleur :|txtColor|T", _
                  "lblChassis|Ch" & ChrW(226) & "ssis :|txtChassis|T", "lblPlate|Immatriculation :|txtPlate|T", _
                  "lblMileage|Kilom" & ChrW(233) & "trage :|txtMileage|T")
    vRight = Array("lblSeller|Vendeur :|cboSeller|C", "lblPrice|Prix d'achat :|txtPrice|T", _
                   "lblIntermediary|Interm" & ChrW(233) & "diaire :|cboIntermediary|C", "lblCommission|Commission :|txtCommission|T", _
                   "lblPlace|Lieu :|cboPlace|C", "lblDate|Date :|txtDate|T")
    For i = LBound(vLeft) To UBound(vLeft)
        AddRow objForm, CStr(vLeft(i)), MARGIN, sngTop + i * ROW_HEIGHT, COL_LABEL, COL_FIELD
    Next i
    For i = LBound(vRight) To UBound(vRight)
        AddRow objForm, CStr(vRight(i)), sngRight, sngTop + i * ROW_HEIGHT, COL_LABEL, COL_FIELD
    Next i

    sngTop = sngTop + (UBound(vLeft) + 1) * ROW_HEIGHT
    AddMessage objForm, sngWidth, sngTop
    sngTop = sngTop + 34
    AddButtons objForm, "Enregistrer", "Annuler", sngWidth, sngTop
    SizeForm objForm, sngWidth, sngTop + BUTTON_HEIGHT + MARGIN
    SetFormCode objForm, CodeFrmCarEdit()
End Sub

Private Sub BuildSalesForm(ByVal objProject As Object)
    Const LIST_WIDTH As Single = 640
    Const LIST_HEIGHT As Single = 230
    Const SIDE_WIDTH As Single = 120
    Dim objForm As Object
    Dim ctl As Object
    Dim sngWidth As Single
    Dim sngTop As Single
    Dim sngLeft As Single
    Dim vHeads As Variant
    Dim vWidths As Variant
    Dim i As Long

    sngWidth = MARGIN * 3 + LIST_WIDTH + SIDE_WIDTH
    Set objForm = NewForm(objProject, "frmSales", "Kamerix - Ventes")
    AddTitle objForm, "lblTitle", "Ventes comptant", sngWidth, MARGIN

    sngTop = MARGIN + 30
    AddLabelAt objForm, "lblStatus", "Statut :", MARGIN, sngTop + 3, 45
    Set ctl = AddControl(objForm, "Forms.ComboBox.1", "cboStatus", MARGIN + 45, sngTop, 110, 20)
    AddLabelAt objForm, "lblCount", "", MARGIN + 170, sngTop + 3, 150

    ' En-tetes (memes largeurs que lstSales.ColumnWidths)
    sngTop = sngTop + 30
    vHeads = Array("Date", "Client", "Voiture", "Prix", "Encaiss" & ChrW(233), "Solde", "Vente", "CV")
    vWidths = Array(55, 110, 150, 65, 65, 65, 55, 55)
    sngLeft = MARGIN + 3
    For i = LBound(vHeads) To UBound(vHeads)
        AddLabelAt objForm, "lblCol" & (i + 1), CStr(vHeads(i)), sngLeft, sngTop, CSng(vWidths(i))
        sngLeft = sngLeft + vWidths(i)
    Next i
    sngTop = sngTop + 16
    Set ctl = AddControl(objForm, "Forms.ListBox.1", "lstSales", MARGIN, sngTop, LIST_WIDTH, LIST_HEIGHT)

    sngLeft = MARGIN * 2 + LIST_WIDTH
    AddButtonAt objForm, "cmdNew", "Nouvelle vente...", sngLeft, sngTop, SIDE_WIDTH
    AddButtonAt objForm, "cmdPay", "Encaisser...", sngLeft, sngTop + BUTTON_HEIGHT + 6, SIDE_WIDTH
    AddButtonAt objForm, "cmdCv", "Suivi du CV...", sngLeft, sngTop + (BUTTON_HEIGHT + 6) * 2, SIDE_WIDTH
    AddButtonAt objForm, "cmdCancel", "Annuler la vente...", sngLeft, sngTop + (BUTTON_HEIGHT + 6) * 3, SIDE_WIDTH
    AddButtonAt objForm, "cmdDocs", "Documents...", sngLeft, sngTop + (BUTTON_HEIGHT + 6) * 4, SIDE_WIDTH

    sngTop = sngTop + LIST_HEIGHT + 6
    AddMessage objForm, sngWidth, sngTop
    sngTop = sngTop + 34
    AddCloseButton objForm, sngWidth, sngTop
    SizeForm objForm, sngWidth, sngTop + BUTTON_HEIGHT + MARGIN
    SetFormCode objForm, CodeFrmSales()
End Sub

Private Sub BuildSaleStartForm(ByVal objProject As Object)
    Const COL_LABEL As Single = 90
    Const COL_FIELD As Single = 270
    Dim objForm As Object
    Dim sngWidth As Single
    Dim sngTop As Single
    Dim i As Long
    Dim vRows As Variant

    sngWidth = MARGIN * 2 + COL_LABEL + COL_FIELD
    Set objForm = NewForm(objProject, "frmSaleStart", "Kamerix - Nouvelle vente")
    AddTitle objForm, "lblTitle", "Nouvelle vente comptant", sngWidth, MARGIN
    sngTop = MARGIN + 34
    vRows = Array("lblCar|Voiture :|cboCar|C", "lblBuyer|Acheteur :|cboBuyer|C", "lblPrice|Prix (TTC) :|txtPrice|T", _
                  "lblIntermediary|Interm" & ChrW(233) & "diaire :|cboIntermediary|C", "lblCommission|Commission :|txtCommission|T", _
                  "lblPlace|Lieu :|cboPlace|C", "lblDate|Date :|txtDate|T")
    For i = LBound(vRows) To UBound(vRows)
        AddRow objForm, CStr(vRows(i)), MARGIN, sngTop + i * ROW_HEIGHT, COL_LABEL, COL_FIELD
    Next i
    sngTop = sngTop + (UBound(vRows) + 1) * ROW_HEIGHT
    AddLabelAt objForm, "lblCost", "", MARGIN + COL_LABEL, sngTop, COL_FIELD
    AddLabelAt objForm, "lblMargin", "", MARGIN + COL_LABEL, sngTop + 18, COL_FIELD
    sngTop = sngTop + 40
    AddMessage objForm, sngWidth, sngTop
    sngTop = sngTop + 34
    AddButtons objForm, "Enregistrer", "Annuler", sngWidth, sngTop
    SizeForm objForm, sngWidth, sngTop + BUTTON_HEIGHT + MARGIN
    SetFormCode objForm, CodeFrmSaleStart()
End Sub

Private Sub BuildSalePaymentForm(ByVal objProject As Object)
    Dim objForm As Object
    Dim ctl As Object
    Dim sngWidth As Single
    Dim sngTop As Single

    sngWidth = MARGIN * 2 + LABEL_WIDTH + FIELD_WIDTH + 60
    Set objForm = NewForm(objProject, "frmSalePayment", "Kamerix - Encaissement")
    AddTitle objForm, "lblTitle", "Encaisser un paiement", sngWidth, MARGIN
    AddInfo objForm, "lblSale", "", sngWidth, MARGIN + 28, 16
    AddInfo objForm, "lblAmounts", "", sngWidth, MARGIN + 46, 16
    sngTop = MARGIN + 72
    AddField objForm, "lblAmount", "Montant :", "txtAmount", False, sngTop
    sngTop = sngTop + ROW_HEIGHT
    AddLabelAt objForm, "lblMode", "Mode :", MARGIN, sngTop + 3, LABEL_WIDTH
    Set ctl = AddControl(objForm, "Forms.ComboBox.1", "cboMode", MARGIN + LABEL_WIDTH, sngTop, FIELD_WIDTH, 20)
    sngTop = sngTop + ROW_HEIGHT
    AddField objForm, "lblDate", "Date :", "txtDate", False, sngTop
    sngTop = sngTop + ROW_HEIGHT
    AddMessage objForm, sngWidth, sngTop
    sngTop = sngTop + 34
    AddButtons objForm, "Encaisser", "Annuler", sngWidth, sngTop
    SizeForm objForm, sngWidth, sngTop + BUTTON_HEIGHT + MARGIN
    SetFormCode objForm, CodeFrmSalePayment()
End Sub

Private Sub BuildSaleCvForm(ByVal objProject As Object)
    Dim objForm As Object
    Dim ctl As Object
    Dim sngWidth As Single
    Dim sngTop As Single

    sngWidth = MARGIN * 2 + LABEL_WIDTH + FIELD_WIDTH + 60
    Set objForm = NewForm(objProject, "frmSaleCv", "Kamerix - Certificat de vente")
    AddTitle objForm, "lblTitle", "Suivi du certificat de vente", sngWidth, MARGIN
    AddInfo objForm, "lblSale", "", sngWidth, MARGIN + 28, 16
    AddInfo objForm, "lblCvStatus", "", sngWidth, MARGIN + 46, 16

    sngTop = MARGIN + 70
    AddSection objForm, "lblSectionSubmit", "D" & ChrW(233) & "p" & ChrW(244) & "t du dossier", MARGIN, sngTop, sngWidth - MARGIN * 2
    sngTop = sngTop + 22
    AddLabelAt objForm, "lblStation", "Commissariat :", MARGIN, sngTop + 3, LABEL_WIDTH
    Set ctl = AddControl(objForm, "Forms.ComboBox.1", "cboStation", MARGIN + LABEL_WIDTH, sngTop, FIELD_WIDTH, 20)
    sngTop = sngTop + ROW_HEIGHT
    AddField objForm, "lblFilingDate", "Date de d" & ChrW(233) & "p" & ChrW(244) & "t :", "txtFilingDate", False, sngTop
    sngTop = sngTop + ROW_HEIGHT
    AddField objForm, "lblFees", "Frais pay" & ChrW(233) & "s :", "txtFees", False, sngTop
    AddButtonAt objForm, "cmdSubmit", "D" & ChrW(233) & "poser", MARGIN + LABEL_WIDTH + FIELD_WIDTH + 6, sngTop - ROW_HEIGHT, 54

    sngTop = sngTop + ROW_HEIGHT + 6
    AddSection objForm, "lblSectionSign", "Signature", MARGIN, sngTop, sngWidth - MARGIN * 2
    sngTop = sngTop + 22
    AddField objForm, "lblCvNumber", "Num" & ChrW(233) & "ro du CV :", "txtCvNumber", False, sngTop
    sngTop = sngTop + ROW_HEIGHT
    AddField objForm, "lblSignatureDate", "Date de signature :", "txtSignatureDate", False, sngTop
    AddButtonAt objForm, "cmdSign", "Sign" & ChrW(233), MARGIN + LABEL_WIDTH + FIELD_WIDTH + 6, sngTop - ROW_HEIGHT / 2, 54

    sngTop = sngTop + ROW_HEIGHT
    AddMessage objForm, sngWidth, sngTop
    sngTop = sngTop + 34
    AddCloseButton objForm, sngWidth, sngTop
    SizeForm objForm, sngWidth, sngTop + BUTTON_HEIGHT + MARGIN
    SetFormCode objForm, CodeFrmSaleCv()
End Sub

Private Sub BuildSaleCancelForm(ByVal objProject As Object)
    Dim objForm As Object
    Dim ctl As Object
    Dim sngWidth As Single
    Dim sngTop As Single

    sngWidth = MARGIN * 2 + LABEL_WIDTH + FIELD_WIDTH + 60
    Set objForm = NewForm(objProject, "frmSaleCancel", "Kamerix - Annulation")
    AddTitle objForm, "lblTitle", "Annuler la vente", sngWidth, MARGIN
    AddInfo objForm, "lblSale", "", sngWidth, MARGIN + 28, 16
    sngTop = MARGIN + 54
    AddField objForm, "lblReason", "Motif :", "txtReason", False, sngTop
    sngTop = sngTop + ROW_HEIGHT
    AddCheckBox objForm, "chkWithdrawal", "D" & ChrW(233) & "sistement de l'acheteur (frais du CV " & ChrW(224) & " sa charge)", _
                MARGIN + LABEL_WIDTH, sngTop, FIELD_WIDTH + 60
    sngTop = sngTop + ROW_HEIGHT
    AddInfo objForm, "lblRefund", "", sngWidth, sngTop, 16
    sngTop = sngTop + 22
    AddLabelAt objForm, "lblMode", "Mode :", MARGIN, sngTop + 3, LABEL_WIDTH
    Set ctl = AddControl(objForm, "Forms.ComboBox.1", "cboMode", MARGIN + LABEL_WIDTH, sngTop, FIELD_WIDTH, 20)
    sngTop = sngTop + ROW_HEIGHT
    AddMessage objForm, sngWidth, sngTop
    sngTop = sngTop + 34
    AddButtons objForm, "Annuler la vente", "Fermer", sngWidth, sngTop
    SizeForm objForm, sngWidth, sngTop + BUTTON_HEIGHT + MARGIN
    SetFormCode objForm, CodeFrmSaleCancel()
End Sub

Private Sub BuildSaleDocumentsForm(ByVal objProject As Object)
    Const LIST_WIDTH As Single = 460
    Const LIST_HEIGHT As Single = 150
    Const SIDE_WIDTH As Single = 100
    Dim objForm As Object
    Dim ctl As Object
    Dim sngWidth As Single
    Dim sngTop As Single
    Dim sngLeft As Single
    Dim vHeads As Variant
    Dim vWidths As Variant
    Dim i As Long

    sngWidth = MARGIN * 3 + LIST_WIDTH + SIDE_WIDTH
    Set objForm = NewForm(objProject, "frmSaleDocuments", "Kamerix - Documents de la vente")
    AddTitle objForm, "lblTitle", "Documents de la vente", sngWidth, MARGIN
    AddInfo objForm, "lblSale", "", sngWidth, MARGIN + 28, 16
    AddLabelAt objForm, "lblCount", "", MARGIN, MARGIN + 46, 150

    ' En-tetes (memes largeurs que lstDocs.ColumnWidths)
    sngTop = MARGIN + 68
    vHeads = Array("Type", "Num" & ChrW(233) & "ro", "Date", "Montant", ChrW(201) & "ditions", "Remarque")
    vWidths = Array(80, 105, 60, 80, 50, 70)
    sngLeft = MARGIN + 3
    For i = LBound(vHeads) To UBound(vHeads)
        AddLabelAt objForm, "lblCol" & (i + 1), CStr(vHeads(i)), sngLeft, sngTop, CSng(vWidths(i))
        sngLeft = sngLeft + vWidths(i)
    Next i
    sngTop = sngTop + 16
    Set ctl = AddControl(objForm, "Forms.ListBox.1", "lstDocs", MARGIN, sngTop, LIST_WIDTH, LIST_HEIGHT)

    sngLeft = MARGIN * 2 + LIST_WIDTH
    AddButtonAt objForm, "cmdPrint", "Imprimer", sngLeft, sngTop, SIDE_WIDTH
    AddButtonAt objForm, "cmdPdf", "PDF", sngLeft, sngTop + BUTTON_HEIGHT + 6, SIDE_WIDTH

    sngTop = sngTop + LIST_HEIGHT + 6
    AddMessage objForm, sngWidth, sngTop
    sngTop = sngTop + 34
    AddCloseButton objForm, sngWidth, sngTop
    SizeForm objForm, sngWidth, sngTop + BUTTON_HEIGHT + MARGIN
    SetFormCode objForm, CodeFrmSaleDocuments()
End Sub

Private Sub BuildSettingsForm(ByVal objProject As Object)
    Const PAGES_WIDTH As Single = 430
    Const PAGES_HEIGHT As Single = 305
    Const IN_LABEL As Single = 140
    Const IN_FIELD As Single = 260
    Const IN_LEFT As Single = 10
    Dim objForm As Object
    Dim mpg As Object
    Dim pg As Object
    Dim ctl As Object
    Dim sngWidth As Single
    Dim sngTop As Single
    Dim i As Long
    Dim vRows As Variant
    Dim vParts As Variant

    sngWidth = MARGIN * 2 + PAGES_WIDTH
    Set objForm = NewForm(objProject, "frmSettings", "Kamerix - Param" & ChrW(232) & "tres")
    AddTitle objForm, "lblTitle", "Param" & ChrW(232) & "tres de Kamerix", sngWidth - 90, MARGIN
    AddInfo objForm, "lblIntro", "", sngWidth - 90, MARGIN + 26, 28
    ' Logo de Kamerix (Images\AppLogo\Logo.png), charge par le code du formulaire
    Set ctl = AddControlIn(objForm.Designer, "Forms.Image.1", "imgAppLogo", sngWidth - MARGIN - 80, MARGIN, 80, 50)
    ctl.BorderStyle = 0
    ctl.BackStyle = 0
    ctl.PictureSizeMode = 3

    sngTop = MARGIN + 60
    Set mpg = AddControl(objForm, "Forms.MultiPage.1", "mpgSettings", MARGIN, sngTop, PAGES_WIDTH, PAGES_HEIGHT)
    mpg.Pages(0).Name = "pgCompany"
    mpg.Pages(0).Caption = "Entreprise"
    mpg.Pages(1).Name = "pgSale"
    mpg.Pages(1).Caption = "Vente"
    mpg.Pages.Add "pgSecurity", "S" & ChrW(233) & "curit" & ChrW(233)

    ' ---------- Onglet Entreprise ----------
    Set pg = mpg.Pages(0)
    vRows = Array("lblName|Nom de l'entreprise * :|txtName", "lblAddress|Adresse * :|txtAddress", _
                  "lblPhone|T" & ChrW(233) & "l" & ChrW(233) & "phone * :|txtPhone", "lblEmail|E-mail :|txtEmail", _
                  "lblNiu|NIU :|txtNiu", "lblRccm|RCCM :|txtRccm", "lblFooter|Pied de page :|txtFooter")
    For i = LBound(vRows) To UBound(vRows)
        vParts = Split(CStr(vRows(i)), "|")
        AddPageRow pg, CStr(vParts(0)), CStr(vParts(1)), CStr(vParts(2)), IN_LEFT, 10 + i * ROW_HEIGHT, IN_LABEL, IN_FIELD, 20
    Next i
    sngTop = 10 + (UBound(vRows) + 1) * ROW_HEIGHT
    Set ctl = AddControlIn(pg, "Forms.Label.1", "lblLogo", IN_LEFT, sngTop + 3, IN_LABEL, 16)
    ctl.Caption = "Logo :"
    Set ctl = AddControlIn(pg, "Forms.Image.1", "imgLogo", IN_LEFT + IN_LABEL, sngTop, 110, 60)
    ctl.BorderStyle = 1
    ctl.PictureSizeMode = 3
    Set ctl = AddControlIn(pg, "Forms.CommandButton.1", "cmdBrowse", IN_LEFT + IN_LABEL + 120, sngTop, 90, BUTTON_HEIGHT)
    ctl.Caption = "Parcourir..."
    Set ctl = AddControlIn(pg, "Forms.CommandButton.1", "cmdRemoveLogo", IN_LEFT + IN_LABEL + 120, sngTop + BUTTON_HEIGHT + 6, 90, BUTTON_HEIGHT)
    ctl.Caption = "Retirer"
    Set ctl = AddControlIn(pg, "Forms.Label.1", "lblLogoInfo", IN_LEFT + IN_LABEL, sngTop + 64, IN_FIELD, 16)
    ctl.Caption = ""
    ctl.ForeColor = RGB(89, 89, 89)

    ' ---------- Onglet Vente ----------
    Set pg = mpg.Pages(1)
    AddPageRow pg, "lblTaxRate", "Taux de TVA (%) :", "txtTaxRate", IN_LEFT, 10, IN_LABEL, 80, 20
    AddPageRow pg, "lblModes", "Modes de paiement (un par ligne) :", "txtModes", IN_LEFT, 10 + ROW_HEIGHT, IN_LABEL, IN_FIELD, 76
    sngTop = 10 + ROW_HEIGHT + 84
    AddPageRow pg, "lblCurrency", "Devise :", "txtCurrency", IN_LEFT, sngTop, IN_LABEL, 80, 20
    AddPageRow pg, "lblDecimals", "D" & ChrW(233) & "cimales :", "txtDecimals", IN_LEFT, sngTop + ROW_HEIGHT, IN_LABEL, 80, 20
    Set ctl = AddControlIn(pg, "Forms.Label.1", "lblLockInfo", IN_LEFT, sngTop + ROW_HEIGHT * 2 + 4, IN_LABEL + IN_FIELD, 30)
    ctl.Caption = ""
    ctl.WordWrap = True
    ctl.ForeColor = RGB(192, 96, 0)
    Set ctl = AddControlIn(pg, "Forms.Label.1", "lblSaleInfo", IN_LEFT, sngTop + ROW_HEIGHT * 2 + 38, IN_LABEL + IN_FIELD, 42)
    ctl.Caption = "Un nouveau taux de TVA s'applique aux prochaines factures ; les factures d" & ChrW(233) & "j" & ChrW(224) & _
                  " " & ChrW(233) & "mises gardent le leur. Devise et d" & ChrW(233) & "cimales se verrouillent d" & ChrW(232) & _
                  "s la premi" & ChrW(232) & "re vente."
    ctl.WordWrap = True
    ctl.ForeColor = RGB(89, 89, 89)

    ' ---------- Onglet Securite ----------
    Set pg = mpg.Pages(2)
    AddPageRow pg, "lblMaxFailures", ChrW(201) & "checs avant blocage :", "txtMaxFailures", IN_LEFT, 10, IN_LABEL, 80, 20
    AddPageRow pg, "lblLockMinutes", "Dur" & ChrW(233) & "e du blocage (minutes) :", "txtLockMinutes", IN_LEFT, 10 + ROW_HEIGHT, _
               IN_LABEL, 80, 20
    Set ctl = AddControlIn(pg, "Forms.Label.1", "lblSecurityInfo", IN_LEFT, 10 + ROW_HEIGHT * 2 + 6, IN_LABEL + IN_FIELD, 42)
    ctl.Caption = "Apr" & ChrW(232) & "s ce nombre d'" & ChrW(233) & "checs de connexion cons" & ChrW(233) & "cutifs, le compte est " & _
                  "bloqu" & ChrW(233) & " pendant la dur" & ChrW(233) & "e indiqu" & ChrW(233) & "e."
    ctl.WordWrap = True
    ctl.ForeColor = RGB(89, 89, 89)

    ' ---------- Bas du formulaire ----------
    sngTop = MARGIN + 60 + PAGES_HEIGHT + 6
    AddMessage objForm, sngWidth, sngTop
    sngTop = sngTop + 34
    AddButtons objForm, "Enregistrer", "Fermer", sngWidth, sngTop
    SizeForm objForm, sngWidth, sngTop + BUTTON_HEIGHT + MARGIN
    SetFormCode objForm, CodeFrmSettings()
End Sub

'==============================================================================
' OUTILS DE CONSTRUCTION
'==============================================================================

' Cree un formulaire vide ; un formulaire du meme nom est sauvegarde puis supprime
Private Function NewForm(ByVal objProject As Object, ByVal strName As String, ByVal strCaption As String) As Object
    Dim objOld As Object
    Dim objForm As Object

    Set objOld = FindComponent(objProject, strName)
    If Not objOld Is Nothing Then
        If Len(mstrBackupFolder) = 0 Then mstrBackupFolder = CreateBackupFolder()
        objOld.Export mstrBackupFolder & Application.PathSeparator & strName & ".frm"
        ' Renommer avant de supprimer : la suppression n'est effective qu'a la fin de la macro
        objOld.Name = TemporaryName(objProject, strName)
        objProject.VBComponents.Remove objOld
    End If

    Set objForm = objProject.VBComponents.Add(VBEXT_MS_FORM)
    objForm.Name = strName
    objForm.Properties("Caption").Value = strCaption
    ' Police du formulaire : heritee par les controles ajoutes ensuite
    objForm.Designer.Font.Name = FONT_NAME
    objForm.Designer.Font.Size = FONT_SIZE
    Set NewForm = objForm
End Function

Private Sub SizeForm(ByVal objForm As Object, ByVal sngWidth As Single, ByVal sngInsideHeight As Single)
    ApplyFonts objForm
    objForm.Properties("Width").Value = sngWidth + 6
    objForm.Properties("Height").Value = sngInsideHeight + FORM_EXTRA_HEIGHT
End Sub

' Repasse la police sur tous les controles (sauf titre et sections) : un reglage fait a la
' creation d'un controle n'est pas toujours conserve par le concepteur MSForms.
Private Sub ApplyFonts(ByVal objForm As Object)
    Dim ctl As Object

    On Error Resume Next                ' certains controles (Image) n'ont pas de police
    For Each ctl In objForm.Designer.Controls
        If ctl.Name <> "lblTitle" And Left$(ctl.Name, 10) <> "lblSection" Then
            ctl.Font.Name = FONT_NAME
            ctl.Font.Size = FONT_SIZE
            ctl.Font.Bold = False
        End If
    Next ctl
    On Error GoTo 0
End Sub

Private Function AddControl(ByVal objForm As Object, ByVal strProgId As String, ByVal strName As String, _
                            ByVal sngLeft As Single, ByVal sngTop As Single, _
                            ByVal sngWidth As Single, ByVal sngHeight As Single) As Object
    Dim ctl As Object

    Set ctl = objForm.Designer.Controls.Add(strProgId, strName, True)
    ctl.Left = sngLeft
    ctl.Top = sngTop
    ctl.Width = sngWidth
    ctl.Height = sngHeight
    ctl.Font.Name = FONT_NAME
    ctl.Font.Size = FONT_SIZE
    Set AddControl = ctl
End Function

Private Sub AddTitle(ByVal objForm As Object, ByVal strName As String, ByVal strCaption As String, _
                     ByVal sngFormWidth As Single, ByVal sngTop As Single)
    Dim ctl As Object

    Set ctl = AddControl(objForm, "Forms.Label.1", strName, MARGIN, sngTop, sngFormWidth - MARGIN * 2, 22)
    ctl.Caption = strCaption
    ctl.Font.Size = 12
    ctl.Font.Bold = True
End Sub

Private Sub AddInfo(ByVal objForm As Object, ByVal strName As String, ByVal strCaption As String, _
                    ByVal sngFormWidth As Single, ByVal sngTop As Single, ByVal sngHeight As Single)
    Dim ctl As Object

    Set ctl = AddControl(objForm, "Forms.Label.1", strName, MARGIN, sngTop, sngFormWidth - MARGIN * 2, sngHeight)
    ctl.Caption = strCaption
    ctl.WordWrap = True
    ctl.ForeColor = RGB(89, 89, 89)
End Sub

' Une ligne de saisie : libelle + zone de texte (masquee pour un mot de passe)
Private Sub AddField(ByVal objForm As Object, ByVal strLabelName As String, ByVal strCaption As String, _
                     ByVal strTextName As String, ByVal blnPassword As Boolean, ByVal sngTop As Single)
    Dim ctl As Object

    Set ctl = AddControl(objForm, "Forms.Label.1", strLabelName, MARGIN, sngTop + 3, LABEL_WIDTH, 16)
    ctl.Caption = strCaption

    Set ctl = AddControl(objForm, "Forms.TextBox.1", strTextName, MARGIN + LABEL_WIDTH, sngTop, FIELD_WIDTH, 20)
    If blnPassword Then ctl.PasswordChar = "*"
End Sub

' Zone des messages d'erreur (rouge, sur deux lignes)
Private Sub AddMessage(ByVal objForm As Object, ByVal sngFormWidth As Single, ByVal sngTop As Single)
    Dim ctl As Object

    Set ctl = AddControl(objForm, "Forms.Label.1", "lblMessage", MARGIN, sngTop, sngFormWidth - MARGIN * 2, 30)
    ctl.Caption = ""
    ctl.WordWrap = True
    ctl.ForeColor = RGB(192, 0, 0)
End Sub

' Boutons alignes a droite : validation (touche Entree) et annulation (touche Echap)
Private Sub AddButtons(ByVal objForm As Object, ByVal strOkCaption As String, ByVal strCancelCaption As String, _
                       ByVal sngFormWidth As Single, ByVal sngTop As Single)
    Dim ctl As Object

    Set ctl = AddControl(objForm, "Forms.CommandButton.1", "cmdOK", _
                         sngFormWidth - MARGIN - BUTTON_WIDTH * 2 - 6, sngTop, BUTTON_WIDTH, BUTTON_HEIGHT)
    ctl.Caption = strOkCaption
    ctl.Default = True

    Set ctl = AddControl(objForm, "Forms.CommandButton.1", "cmdCancel", _
                         sngFormWidth - MARGIN - BUTTON_WIDTH, sngTop, BUTTON_WIDTH, BUTTON_HEIGHT)
    ctl.Caption = strCancelCaption
    ctl.Cancel = True
End Sub

' Libelle simple a une position donnee
Private Sub AddLabelAt(ByVal objForm As Object, ByVal strName As String, ByVal strCaption As String, _
                       ByVal sngLeft As Single, ByVal sngTop As Single, ByVal sngWidth As Single)
    Dim ctl As Object

    Set ctl = AddControl(objForm, "Forms.Label.1", strName, sngLeft, sngTop, sngWidth, 16)
    ctl.Caption = strCaption
End Sub

' Libelle cliquable presente comme un lien (bleu, souligne)
Private Sub AddLink(ByVal objForm As Object, ByVal strName As String, ByVal strCaption As String, _
                    ByVal sngLeft As Single, ByVal sngTop As Single, ByVal sngWidth As Single)
    Dim ctl As Object

    Set ctl = AddControl(objForm, "Forms.Label.1", strName, sngLeft, sngTop, sngWidth, 16)
    ctl.Caption = strCaption
    ctl.ForeColor = RGB(0, 102, 204)
    ctl.Font.Underline = True
End Sub

' Bouton a une position donnee
Private Sub AddButtonAt(ByVal objForm As Object, ByVal strName As String, ByVal strCaption As String, _
                        ByVal sngLeft As Single, ByVal sngTop As Single, ByVal sngWidth As Single)
    Dim ctl As Object

    Set ctl = AddControl(objForm, "Forms.CommandButton.1", strName, sngLeft, sngTop, sngWidth, BUTTON_HEIGHT)
    ctl.Caption = strCaption
End Sub

' Case a cocher
Private Sub AddCheckBox(ByVal objForm As Object, ByVal strName As String, ByVal strCaption As String, _
                        ByVal sngLeft As Single, ByVal sngTop As Single, ByVal sngWidth As Single)
    Dim ctl As Object

    Set ctl = AddControl(objForm, "Forms.CheckBox.1", strName, sngLeft, sngTop, sngWidth, 18)
    ctl.Caption = strCaption
End Sub

' Titre de section (gras, bleu) ; son nom doit commencer par lblSection (ApplyFonts le respecte)
Private Sub AddSection(ByVal objForm As Object, ByVal strName As String, ByVal strCaption As String, _
                       ByVal sngLeft As Single, ByVal sngTop As Single, ByVal sngWidth As Single)
    Dim ctl As Object

    Set ctl = AddControl(objForm, "Forms.Label.1", strName, sngLeft, sngTop, sngWidth, 18)
    ctl.Caption = strCaption
    ctl.Font.Bold = True
    ctl.ForeColor = RGB(46, 116, 181)
End Sub

' Ligne libelle + champ decrite par "nomLibelle|texte|nomControle|C ou T"
Private Sub AddRow(ByVal objForm As Object, ByVal strSpec As String, ByVal sngLeft As Single, ByVal sngTop As Single, _
                   ByVal sngLabelWidth As Single, ByVal sngFieldWidth As Single)
    Dim vParts As Variant
    Dim ctl As Object

    vParts = Split(strSpec, "|")
    AddLabelAt objForm, CStr(vParts(0)), CStr(vParts(1)), sngLeft, sngTop + 3, sngLabelWidth
    If vParts(3) = "C" Then
        Set ctl = AddControl(objForm, "Forms.ComboBox.1", CStr(vParts(2)), sngLeft + sngLabelWidth, sngTop, sngFieldWidth, 20)
    Else
        Set ctl = AddControl(objForm, "Forms.TextBox.1", CStr(vParts(2)), sngLeft + sngLabelWidth, sngTop, sngFieldWidth, 20)
    End If
End Sub

' Bouton Fermer (touche Echap), en bas a droite
Private Sub AddCloseButton(ByVal objForm As Object, ByVal sngFormWidth As Single, ByVal sngTop As Single)
    Dim ctl As Object

    Set ctl = AddControl(objForm, "Forms.CommandButton.1", "cmdClose", sngFormWidth - MARGIN - BUTTON_WIDTH, sngTop, _
                         BUTTON_WIDTH, BUTTON_HEIGHT)
    ctl.Caption = "Fermer"
    ctl.Cancel = True
End Sub

' Controle dans un conteneur (page d'un MultiPage, cadre) ; le controle Image n'a pas de police
Private Function AddControlIn(ByVal objContainer As Object, ByVal strProgId As String, ByVal strName As String, _
                              ByVal sngLeft As Single, ByVal sngTop As Single, _
                              ByVal sngWidth As Single, ByVal sngHeight As Single) As Object
    Dim ctl As Object

    Set ctl = objContainer.Controls.Add(strProgId, strName, True)
    ctl.Left = sngLeft
    ctl.Top = sngTop
    ctl.Width = sngWidth
    ctl.Height = sngHeight
    On Error Resume Next
    ctl.Font.Name = FONT_NAME
    ctl.Font.Size = FONT_SIZE
    On Error GoTo 0
    Set AddControlIn = ctl
End Function

' Ligne libelle + zone de texte dans un conteneur
Private Sub AddPageRow(ByVal objContainer As Object, ByVal strLabelName As String, ByVal strCaption As String, _
                       ByVal strTextName As String, ByVal sngLeft As Single, ByVal sngTop As Single, _
                       ByVal sngLabelWidth As Single, ByVal sngFieldWidth As Single, ByVal sngFieldHeight As Single)
    Dim ctl As Object

    Set ctl = AddControlIn(objContainer, "Forms.Label.1", strLabelName, sngLeft, sngTop + 3, sngLabelWidth, 30)
    ctl.Caption = strCaption
    ctl.WordWrap = True
    Set ctl = AddControlIn(objContainer, "Forms.TextBox.1", strTextName, sngLeft + sngLabelWidth, sngTop, sngFieldWidth, sngFieldHeight)
End Sub

' Remplace tout le code du formulaire
Private Sub SetFormCode(ByVal objForm As Object, ByVal strCode As String)
    With objForm.CodeModule
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        .AddFromString strCode
    End With
End Sub

'==============================================================================
' OUTILS DU PROJET
'==============================================================================

Private Function GetProject() As Object
    Dim objProject As Object

    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox "Enregistrez d'abord le classeur.", vbExclamation, TOOL_TITLE
        Exit Function
    End If

    On Error Resume Next
    Set objProject = ThisWorkbook.VBProject
    If Err.Number <> 0 Or objProject Is Nothing Then
        On Error GoTo 0
        MsgBox "Acces au projet VBA refuse." & vbCrLf & _
               "Cochez 'Acces approuve au modele d'objet du projet VBA' dans le " & _
               "Centre de gestion de la confidentialite, puis relancez.", vbExclamation, TOOL_TITLE
        Exit Function
    End If
    On Error GoTo 0

    If objProject.Protection = VBEXT_PP_LOCKED Then
        MsgBox "Le projet VBA est protege par mot de passe : deverrouillez-le d'abord.", _
               vbExclamation, TOOL_TITLE
        Exit Function
    End If

    Set GetProject = objProject
End Function

Private Function FindComponent(ByVal objProject As Object, ByVal strName As String) As Object
    Dim objComp As Object

    For Each objComp In objProject.VBComponents
        If StrComp(objComp.Name, strName, vbTextCompare) = 0 Then
            Set FindComponent = objComp
            Exit Function
        End If
    Next objComp
End Function

Private Function TemporaryName(ByVal objProject As Object, ByVal strName As String) As String
    Dim strCandidate As String
    Dim i As Long

    Do
        i = i + 1
        strCandidate = Left$(strName, MAX_MODULE_NAME - 6) & "_old" & i
    Loop Until FindComponent(objProject, strCandidate) Is Nothing

    TemporaryName = strCandidate
End Function

Private Function CreateBackupFolder() As String
    Dim strFolder As String

    strFolder = ThisWorkbook.Path & Application.PathSeparator & _
                "Backup_VBA_" & Format$(Now, "yyyymmdd_hhnnss")
    If Len(Dir$(strFolder, vbDirectory)) = 0 Then MkDir strFolder
    CreateBackupFolder = strFolder
End Function

'==============================================================================
' CODE DES FORMULAIRES (genere : une ligne de code par instruction s = s & ...)
'==============================================================================

Private Function CodeFrmLogin() As String
    Dim s As String

    s = s & "Option Explicit" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' KAMERIX - Ecran de connexion (cree par ZmodFormBuilder : ne pas modifier a la main," & vbCrLf
    s = s & "' modifier le generateur puis relancer BuildKamerixForms)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private mblnSucceeded As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Vrai si la session a ete ouverte (lu par modKamerixLogin apres la fermeture)" & vbCrLf
    s = s & "Public Property Get Succeeded() As Boolean" & vbCrLf
    s = s & "    Succeeded = mblnSucceeded" & vbCrLf
    s = s & "End Property" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_Initialize()" & vbCrLf
    s = s & "    lblMessage.Caption = """"" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdOK_Click()" & vbCrLf
    s = s & "    Dim eResult As KxLoginResult" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    lblMessage.Caption = """"" & vbCrLf
    s = s & "    If Len(Trim$(txtLogin.Text)) = 0 Or Len(txtPassword.Text) = 0 Then" & vbCrLf
    s = s & "        lblMessage.Caption = ""Saisissez l'identifiant et le mot de passe.""" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo EH" & vbCrLf
    s = s & "    SetBusy True" & vbCrLf
    s = s & "    eResult = OpenSession(txtLogin.Text, txtPassword.Text)" & vbCrLf
    s = s & "    SetBusy False" & vbCrLf
    s = s & "    txtPassword.Text = """"" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If eResult = kxLoginSuccess Then" & vbCrLf
    s = s & "        mblnSucceeded = True" & vbCrLf
    s = s & "        Me.Hide" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        lblMessage.Caption = LoginResultMessage(eResult)" & vbCrLf
    s = s & "        txtPassword.SetFocus" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "EH:" & vbCrLf
    s = s & "    SetBusy False" & vbCrLf
    s = s & "    txtPassword.Text = """"" & vbCrLf
    s = s & "    lblMessage.Caption = ""Erreur "" & Err.Number & "" : "" & Err.Description" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdCancel_Click()" & vbCrLf
    s = s & "    Me.Hide" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Lien ""Mot de passe oublie ?"" : redefinition par le code de secours (administrateurs)" & vbCrLf
    s = s & "Private Sub lblForgot_Click()" & vbCrLf
    s = s & "    ShowRecoveryForm" & vbCrLf
    s = s & "    txtPassword.Text = """"" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' La croix ferme l'ecran comme Annuler (le formulaire reste charge pour modKamerixLogin)" & vbCrLf
    s = s & "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & vbCrLf
    s = s & "    If CloseMode = 0 Then" & vbCrLf
    s = s & "        Cancel = True" & vbCrLf
    s = s & "        Me.Hide" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Sablier et boutons inactifs pendant le calcul du mot de passe (environ 1 seconde)" & vbCrLf
    s = s & "Private Sub SetBusy(ByVal blnBusy As Boolean)" & vbCrLf
    s = s & "    Me.MousePointer = IIf(blnBusy, 11, 0)" & vbCrLf
    s = s & "    cmdOK.Enabled = Not blnBusy" & vbCrLf
    s = s & "    cmdCancel.Enabled = Not blnBusy" & vbCrLf
    s = s & "    DoEvents" & vbCrLf
    s = s & "End Sub" & vbCrLf
    CodeFrmLogin = s
End Function

Private Function CodeFrmFirstAdmin() As String
    Dim s As String

    s = s & "Option Explicit" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' KAMERIX - Premier lancement : creation de l'administrateur (cree par ZmodFormBuilder :" & vbCrLf
    s = s & "' ne pas modifier a la main, modifier le generateur puis relancer BuildKamerixForms)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private mblnSucceeded As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Vrai si l'administrateur a ete cree et sa session ouverte" & vbCrLf
    s = s & "Public Property Get Succeeded() As Boolean" & vbCrLf
    s = s & "    Succeeded = mblnSucceeded" & vbCrLf
    s = s & "End Property" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_Initialize()" & vbCrLf
    s = s & "    lblRules.Caption = PasswordRulesText()" & vbCrLf
    s = s & "    lblMessage.Caption = """"" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdOK_Click()" & vbCrLf
    s = s & "    Dim strMsg As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    lblMessage.Caption = """"" & vbCrLf
    s = s & "    If txtPassword.Text <> txtConfirm.Text Then" & vbCrLf
    s = s & "        lblMessage.Caption = ""La confirmation ne correspond pas au mot de passe.""" & vbCrLf
    s = s & "        txtConfirm.SetFocus" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    strMsg = ValidatePasswordRules(txtPassword.Text, txtLogin.Text)" & vbCrLf
    s = s & "    If Len(strMsg) > 0 Then" & vbCrLf
    s = s & "        lblMessage.Caption = strMsg" & vbCrLf
    s = s & "        txtPassword.SetFocus" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo EH" & vbCrLf
    s = s & "    SetBusy True" & vbCrLf
    s = s & "    CreateFirstAdmin txtLogin.Text, txtFullName.Text, txtPassword.Text" & vbCrLf
    s = s & "    mblnSucceeded = (OpenSession(txtLogin.Text, txtPassword.Text) = kxLoginSuccess)" & vbCrLf
    s = s & "    SetBusy False" & vbCrLf
    s = s & "    ClearPasswords" & vbCrLf
    s = s & "    Me.Hide" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "EH:" & vbCrLf
    s = s & "    SetBusy False" & vbCrLf
    s = s & "    lblMessage.Caption = Err.Description" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdCancel_Click()" & vbCrLf
    s = s & "    ClearPasswords" & vbCrLf
    s = s & "    Me.Hide" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & vbCrLf
    s = s & "    If CloseMode = 0 Then" & vbCrLf
    s = s & "        Cancel = True" & vbCrLf
    s = s & "        ClearPasswords" & vbCrLf
    s = s & "        Me.Hide" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub ClearPasswords()" & vbCrLf
    s = s & "    txtPassword.Text = """"" & vbCrLf
    s = s & "    txtConfirm.Text = """"" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub SetBusy(ByVal blnBusy As Boolean)" & vbCrLf
    s = s & "    Me.MousePointer = IIf(blnBusy, 11, 0)" & vbCrLf
    s = s & "    cmdOK.Enabled = Not blnBusy" & vbCrLf
    s = s & "    cmdCancel.Enabled = Not blnBusy" & vbCrLf
    s = s & "    DoEvents" & vbCrLf
    s = s & "End Sub" & vbCrLf
    CodeFrmFirstAdmin = s
End Function

Private Function CodeFrmChangePassword() As String
    Dim s As String

    s = s & "Option Explicit" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' KAMERIX - Changement du mot de passe (cree par ZmodFormBuilder : ne pas modifier a la" & vbCrLf
    s = s & "' main, modifier le generateur puis relancer BuildKamerixForms)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private mblnSucceeded As Boolean" & vbCrLf
    s = s & "Private mblnForced As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Vrai si le mot de passe a ete change" & vbCrLf
    s = s & "Public Property Get Succeeded() As Boolean" & vbCrLf
    s = s & "    Succeeded = mblnSucceeded" & vbCrLf
    s = s & "End Property" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Vrai : mot de passe provisoire, changement obligatoire (renoncer ferme la session)" & vbCrLf
    s = s & "Public Property Let Forced(ByVal blnForced As Boolean)" & vbCrLf
    s = s & "    mblnForced = blnForced" & vbCrLf
    s = s & "    If blnForced Then" & vbCrLf
    s = s & "        lblInfo.Caption = ""Votre mot de passe est provisoire : choisissez-en un nouveau pour continuer.""" & vbCrLf
    s = s & "        cmdCancel.Caption = ""Quitter""" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        lblInfo.Caption = ""Saisissez votre mot de passe actuel, puis le nouveau.""" & vbCrLf
    s = s & "        cmdCancel.Caption = ""Annuler""" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Property" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_Initialize()" & vbCrLf
    s = s & "    lblRules.Caption = PasswordRulesText()" & vbCrLf
    s = s & "    lblMessage.Caption = """"" & vbCrLf
    s = s & "    Me.Forced = False" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdOK_Click()" & vbCrLf
    s = s & "    Dim strMsg As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    lblMessage.Caption = """"" & vbCrLf
    s = s & "    If txtNew.Text <> txtConfirm.Text Then" & vbCrLf
    s = s & "        lblMessage.Caption = ""La confirmation ne correspond pas au nouveau mot de passe.""" & vbCrLf
    s = s & "        txtConfirm.SetFocus" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    strMsg = ValidatePasswordRules(txtNew.Text, CurrentUserLogin())" & vbCrLf
    s = s & "    If Len(strMsg) > 0 Then" & vbCrLf
    s = s & "        lblMessage.Caption = strMsg" & vbCrLf
    s = s & "        txtNew.SetFocus" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo EH" & vbCrLf
    s = s & "    SetBusy True" & vbCrLf
    s = s & "    ChangeMyPassword txtOld.Text, txtNew.Text" & vbCrLf
    s = s & "    SetBusy False" & vbCrLf
    s = s & "    mblnSucceeded = True" & vbCrLf
    s = s & "    ClearPasswords" & vbCrLf
    s = s & "    Me.Hide" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "EH:" & vbCrLf
    s = s & "    SetBusy False" & vbCrLf
    s = s & "    lblMessage.Caption = Err.Description" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdCancel_Click()" & vbCrLf
    s = s & "    ClearPasswords" & vbCrLf
    s = s & "    Me.Hide" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & vbCrLf
    s = s & "    If CloseMode = 0 Then" & vbCrLf
    s = s & "        Cancel = True" & vbCrLf
    s = s & "        ClearPasswords" & vbCrLf
    s = s & "        Me.Hide" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub ClearPasswords()" & vbCrLf
    s = s & "    txtOld.Text = """"" & vbCrLf
    s = s & "    txtNew.Text = """"" & vbCrLf
    s = s & "    txtConfirm.Text = """"" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub SetBusy(ByVal blnBusy As Boolean)" & vbCrLf
    s = s & "    Me.MousePointer = IIf(blnBusy, 11, 0)" & vbCrLf
    s = s & "    cmdOK.Enabled = Not blnBusy" & vbCrLf
    s = s & "    cmdCancel.Enabled = Not blnBusy" & vbCrLf
    s = s & "    DoEvents" & vbCrLf
    s = s & "End Sub" & vbCrLf
    CodeFrmChangePassword = s
End Function

Private Function CodeFrmRecovery() As String
    Dim s As String

    s = s & "Option Explicit" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' KAMERIX - Code de secours : redefinir le mot de passe d'un administrateur (cree par" & vbCrLf
    s = s & "' ZmodFormBuilder : ne pas modifier a la main, modifier le generateur puis relancer BuildKamerixForms)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private mblnSucceeded As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Vrai si le mot de passe a ete redefini" & vbCrLf
    s = s & "Public Property Get Succeeded() As Boolean" & vbCrLf
    s = s & "    Succeeded = mblnSucceeded" & vbCrLf
    s = s & "End Property" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_Initialize()" & vbCrLf
    s = s & "    lblRules.Caption = PasswordRulesText()" & vbCrLf
    s = s & "    lblMessage.Caption = """"" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdOK_Click()" & vbCrLf
    s = s & "    Dim strMsg As String" & vbCrLf
    s = s & "    Dim strNewCode As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    lblMessage.Caption = """"" & vbCrLf
    s = s & "    If Len(Trim$(txtCode.Text)) = 0 Or Len(Trim$(txtLogin.Text)) = 0 Then" & vbCrLf
    s = s & "        lblMessage.Caption = ""Saisissez le code de secours et l'identifiant de l'administrateur.""" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If txtNew.Text <> txtConfirm.Text Then" & vbCrLf
    s = s & "        lblMessage.Caption = ""La confirmation ne correspond pas au nouveau mot de passe.""" & vbCrLf
    s = s & "        txtConfirm.SetFocus" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    strMsg = ValidatePasswordRules(txtNew.Text, txtLogin.Text)" & vbCrLf
    s = s & "    If Len(strMsg) > 0 Then" & vbCrLf
    s = s & "        lblMessage.Caption = strMsg" & vbCrLf
    s = s & "        txtNew.SetFocus" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo EH" & vbCrLf
    s = s & "    SetBusy True" & vbCrLf
    s = s & "    strNewCode = ResetAdminWithRecoveryCode(txtCode.Text, txtLogin.Text, txtNew.Text)" & vbCrLf
    s = s & "    SetBusy False" & vbCrLf
    s = s & "    mblnSucceeded = True" & vbCrLf
    s = s & "    ClearFields" & vbCrLf
    s = s & "    Me.Hide" & vbCrLf
    s = s & "    ShowRecoveryCode strNewCode, ""Mot de passe red"" & ChrW(233) & ""fini : connectez-vous avec le nouveau mot de passe."" & vbCrLf & _" & vbCrLf
    s = s & "        ""L'ancien code de secours ne fonctionne plus.""" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "EH:" & vbCrLf
    s = s & "    SetBusy False" & vbCrLf
    s = s & "    lblMessage.Caption = Err.Description" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdCancel_Click()" & vbCrLf
    s = s & "    ClearFields" & vbCrLf
    s = s & "    Me.Hide" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & vbCrLf
    s = s & "    If CloseMode = 0 Then" & vbCrLf
    s = s & "        Cancel = True" & vbCrLf
    s = s & "        ClearFields" & vbCrLf
    s = s & "        Me.Hide" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub ClearFields()" & vbCrLf
    s = s & "    txtCode.Text = """"" & vbCrLf
    s = s & "    txtNew.Text = """"" & vbCrLf
    s = s & "    txtConfirm.Text = """"" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub SetBusy(ByVal blnBusy As Boolean)" & vbCrLf
    s = s & "    Me.MousePointer = IIf(blnBusy, 11, 0)" & vbCrLf
    s = s & "    cmdOK.Enabled = Not blnBusy" & vbCrLf
    s = s & "    cmdCancel.Enabled = Not blnBusy" & vbCrLf
    s = s & "    DoEvents" & vbCrLf
    s = s & "End Sub" & vbCrLf
    CodeFrmRecovery = s
End Function

Private Function CodeFrmUsers() As String
    Dim s As String

    s = s & "Option Explicit" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' KAMERIX - Gestion des comptes (cree par ZmodFormBuilder : ne pas modifier a la main," & vbCrLf
    s = s & "' modifier le generateur puis relancer BuildKamerixForms)" & vbCrLf
    s = s & "' Colonnes de lstUsers : 0 UtilisateurID (masquee), 1 identifiant, 2 nom, 3 role," & vbCrLf
    s = s & "' 4 etat, 5 actif ""1""/""0"" (masquee), 6 RoleID (masquee)." & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Const COL_ID As Long = 0" & vbCrLf
    s = s & "Private Const COL_LOGIN As Long = 1" & vbCrLf
    s = s & "Private Const COL_NAME As Long = 2" & vbCrLf
    s = s & "Private Const COL_STATE As Long = 4" & vbCrLf
    s = s & "Private Const COL_ACTIVE As Long = 5" & vbCrLf
    s = s & "Private Const COL_ROLE_ID As Long = 6" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Toujours Vrai : l'ecran ne fait que consulter et agir (lu par modKamerixLogin)" & vbCrLf
    s = s & "Public Property Get Succeeded() As Boolean" & vbCrLf
    s = s & "    Succeeded = True" & vbCrLf
    s = s & "End Property" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_Initialize()" & vbCrLf
    s = s & "    lstUsers.ColumnCount = 7" & vbCrLf
    s = s & "    lstUsers.ColumnWidths = ""0 pt;80 pt;130 pt;85 pt;90 pt;0 pt;0 pt""" & vbCrLf
    s = s & "    lblMessage.Caption = """"" & vbCrLf
    s = s & "    LoadList 0" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Recharge la liste ; pSelectId = compte a reselectionner (0 = aucun)" & vbCrLf
    s = s & "Private Sub LoadList(ByVal pSelectId As Long)" & vbCrLf
    s = s & "    Dim AdoRs As ADODB.Recordset" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    Dim strState As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    lstUsers.Clear" & vbCrLf
    s = s & "    Set AdoRs = ListAccounts()" & vbCrLf
    s = s & "    Do While Not AdoRs.EOF" & vbCrLf
    s = s & "        If Not CBool(AdoRs.Fields(""Actif"").Value) Then" & vbCrLf
    s = s & "            strState = ""Inactif""" & vbCrLf
    s = s & "        ElseIf Not IsNull(AdoRs.Fields(""DateBlocage"").Value) Then" & vbCrLf
    s = s & "            If CDate(AdoRs.Fields(""DateBlocage"").Value) > Now Then strState = ""Bloqu"" & ChrW(233) Else strState = ""Actif""" & vbCrLf
    s = s & "        ElseIf CBool(AdoRs.Fields(""DoitChangerMdp"").Value) Then" & vbCrLf
    s = s & "            strState = ""Mdp provisoire""" & vbCrLf
    s = s & "        Else" & vbCrLf
    s = s & "            strState = ""Actif""" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "        lstUsers.AddItem CStr(AdoRs.Fields(""UtilisateurID"").Value)" & vbCrLf
    s = s & "        i = lstUsers.ListCount - 1" & vbCrLf
    s = s & "        lstUsers.List(i, COL_LOGIN) = CStr(AdoRs.Fields(""Login"").Value)" & vbCrLf
    s = s & "        lstUsers.List(i, COL_NAME) = CStr(AdoRs.Fields(""NomComplet"").Value)" & vbCrLf
    s = s & "        lstUsers.List(i, 3) = CStr(AdoRs.Fields(""NomRole"").Value)" & vbCrLf
    s = s & "        lstUsers.List(i, COL_STATE) = strState" & vbCrLf
    s = s & "        lstUsers.List(i, COL_ACTIVE) = IIf(CBool(AdoRs.Fields(""Actif"").Value), ""1"", ""0"")" & vbCrLf
    s = s & "        lstUsers.List(i, COL_ROLE_ID) = CStr(AdoRs.Fields(""RoleID"").Value)" & vbCrLf
    s = s & "        If CLng(AdoRs.Fields(""UtilisateurID"").Value) = pSelectId Then lstUsers.ListIndex = i" & vbCrLf
    s = s & "        AdoRs.MoveNext" & vbCrLf
    s = s & "    Loop" & vbCrLf
    s = s & "    AdoRs.Close" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    UpdateButtons" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub lstUsers_Change()" & vbCrLf
    s = s & "    lblMessage.Caption = """"" & vbCrLf
    s = s & "    UpdateButtons" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Boutons actifs selon la selection et le droit d'ecriture sur Administration" & vbCrLf
    s = s & "Private Sub UpdateButtons()" & vbCrLf
    s = s & "    Dim blnWrite As Boolean" & vbCrLf
    s = s & "    Dim blnSelected As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    blnWrite = CanWrite(MOD_ADMINISTRATION)" & vbCrLf
    s = s & "    blnSelected = (lstUsers.ListIndex >= 0)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    cmdNew.Enabled = blnWrite" & vbCrLf
    s = s & "    cmdRecovery.Enabled = blnWrite" & vbCrLf
    s = s & "    cmdRole.Enabled = blnWrite And blnSelected" & vbCrLf
    s = s & "    cmdReset.Enabled = blnWrite And blnSelected" & vbCrLf
    s = s & "    cmdUnlock.Enabled = blnWrite And blnSelected" & vbCrLf
    s = s & "    cmdActive.Enabled = blnWrite And blnSelected" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If blnSelected Then" & vbCrLf
    s = s & "        cmdActive.Caption = IIf(SelectedText(COL_ACTIVE) = ""1"", ""D"" & ChrW(233) & ""sactiver"", ""Activer"")" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        cmdActive.Caption = ""D"" & ChrW(233) & ""sactiver""" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function SelectedId() As Long" & vbCrLf
    s = s & "    SelectedId = CLng(lstUsers.List(lstUsers.ListIndex, COL_ID))" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function SelectedText(ByVal pColumn As Long) As String" & vbCrLf
    s = s & "    SelectedText = CStr(lstUsers.List(lstUsers.ListIndex, pColumn))" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdNew_Click()" & vbCrLf
    s = s & "    Dim frm As Object" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo EH" & vbCrLf
    s = s & "    Set frm = VBA.UserForms.Add(""frmUserEdit"")" & vbCrLf
    s = s & "    frm.PrepareNew" & vbCrLf
    s = s & "    frm.Show" & vbCrLf
    s = s & "    If frm.Succeeded Then LoadList 0" & vbCrLf
    s = s & "    Unload frm" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "EH:" & vbCrLf
    s = s & "    lblMessage.Caption = Err.Description" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdRole_Click()" & vbCrLf
    s = s & "    Dim frm As Object" & vbCrLf
    s = s & "    Dim lngId As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo EH" & vbCrLf
    s = s & "    lngId = SelectedId()" & vbCrLf
    s = s & "    Set frm = VBA.UserForms.Add(""frmUserEdit"")" & vbCrLf
    s = s & "    frm.PrepareRole lngId, SelectedText(COL_LOGIN), SelectedText(COL_NAME), CLng(SelectedText(COL_ROLE_ID))" & vbCrLf
    s = s & "    frm.Show" & vbCrLf
    s = s & "    If frm.Succeeded Then LoadList lngId" & vbCrLf
    s = s & "    Unload frm" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "EH:" & vbCrLf
    s = s & "    lblMessage.Caption = Err.Description" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdReset_Click()" & vbCrLf
    s = s & "    Dim lngId As Long" & vbCrLf
    s = s & "    Dim strLogin As String" & vbCrLf
    s = s & "    Dim strTemp As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo EH" & vbCrLf
    s = s & "    lngId = SelectedId()" & vbCrLf
    s = s & "    strLogin = SelectedText(COL_LOGIN)" & vbCrLf
    s = s & "    If MsgBox(""Donner un nouveau mot de passe provisoire "" & ChrW(224) & "" "" & strLogin & "" ?"" & vbCrLf & _" & vbCrLf
    s = s & "              ""L'ancien mot de passe ne fonctionnera plus."", vbQuestion + vbYesNo, ""Kamerix"") <> vbYes Then Exit Sub" & vbCrLf
    s = s & "    Me.MousePointer = 11" & vbCrLf
    s = s & "    strTemp = ResetAccountWithTemporaryPassword(lngId)" & vbCrLf
    s = s & "    Me.MousePointer = 0" & vbCrLf
    s = s & "    LoadList lngId" & vbCrLf
    s = s & "    ShowTemporaryPassword strLogin, strTemp" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "EH:" & vbCrLf
    s = s & "    Me.MousePointer = 0" & vbCrLf
    s = s & "    lblMessage.Caption = Err.Description" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdUnlock_Click()" & vbCrLf
    s = s & "    Dim lngId As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo EH" & vbCrLf
    s = s & "    lngId = SelectedId()" & vbCrLf
    s = s & "    UnlockAccount lngId" & vbCrLf
    s = s & "    LoadList lngId" & vbCrLf
    s = s & "    lblMessage.Caption = ""Compte d"" & ChrW(233) & ""bloqu"" & ChrW(233) & "" : "" & SelectedText(COL_LOGIN)" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "EH:" & vbCrLf
    s = s & "    lblMessage.Caption = Err.Description" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdActive_Click()" & vbCrLf
    s = s & "    Dim lngId As Long" & vbCrLf
    s = s & "    Dim blnActivate As Boolean" & vbCrLf
    s = s & "    Dim strLogin As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo EH" & vbCrLf
    s = s & "    lngId = SelectedId()" & vbCrLf
    s = s & "    strLogin = SelectedText(COL_LOGIN)" & vbCrLf
    s = s & "    blnActivate = (SelectedText(COL_ACTIVE) = ""0"")" & vbCrLf
    s = s & "    If Not blnActivate Then" & vbCrLf
    s = s & "        If MsgBox(""D"" & ChrW(233) & ""sactiver le compte "" & strLogin & "" ?"" & vbCrLf & _" & vbCrLf
    s = s & "                  ""Il ne pourra plus se connecter."", vbQuestion + vbYesNo, ""Kamerix"") <> vbYes Then Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    SetAccountActive lngId, blnActivate" & vbCrLf
    s = s & "    LoadList lngId" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "EH:" & vbCrLf
    s = s & "    lblMessage.Caption = Err.Description" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdRecovery_Click()" & vbCrLf
    s = s & "    Dim strCode As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo EH" & vbCrLf
    s = s & "    If MsgBox(""Cr"" & ChrW(233) & ""er un nouveau code de secours ?"" & vbCrLf & _" & vbCrLf
    s = s & "              ""L'ancien code ne fonctionnera plus."", vbQuestion + vbYesNo, ""Kamerix"") <> vbYes Then Exit Sub" & vbCrLf
    s = s & "    Me.MousePointer = 11" & vbCrLf
    s = s & "    strCode = RenewRecoveryCode()" & vbCrLf
    s = s & "    Me.MousePointer = 0" & vbCrLf
    s = s & "    ShowRecoveryCode strCode, ""Nouveau code de secours cr"" & ChrW(233) & """" & ChrW(233) & "". L'ancien ne fonctionne plus.""" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "EH:" & vbCrLf
    s = s & "    Me.MousePointer = 0" & vbCrLf
    s = s & "    lblMessage.Caption = Err.Description" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdClose_Click()" & vbCrLf
    s = s & "    Me.Hide" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & vbCrLf
    s = s & "    If CloseMode = 0 Then" & vbCrLf
    s = s & "        Cancel = True" & vbCrLf
    s = s & "        Me.Hide" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    CodeFrmUsers = s
End Function

Private Function CodeFrmUserEdit() As String
    Dim s As String

    s = s & "Option Explicit" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' KAMERIX - Nouveau compte ou changement de role (cree par ZmodFormBuilder : ne pas" & vbCrLf
    s = s & "' modifier a la main, modifier le generateur puis relancer BuildKamerixForms)" & vbCrLf
    s = s & "' Appeler PrepareNew ou PrepareRole avant Show." & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private mblnSucceeded As Boolean" & vbCrLf
    s = s & "Private mblnRoleMode As Boolean" & vbCrLf
    s = s & "Private mlngUserId As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Vrai si le compte a ete cree ou son role change" & vbCrLf
    s = s & "Public Property Get Succeeded() As Boolean" & vbCrLf
    s = s & "    Succeeded = mblnSucceeded" & vbCrLf
    s = s & "End Property" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Creation d'un compte : un mot de passe provisoire sera genere et affiche une fois" & vbCrLf
    s = s & "Public Sub PrepareNew()" & vbCrLf
    s = s & "    mblnRoleMode = False" & vbCrLf
    s = s & "    Me.Caption = ""Kamerix - Nouveau compte""" & vbCrLf
    s = s & "    lblTitle.Caption = ""Nouveau compte""" & vbCrLf
    s = s & "    lblInfo.Caption = ""Un mot de passe provisoire sera g"" & ChrW(233) & ""n"" & ChrW(233) & ""r"" & ChrW(233) & _" & vbCrLf
    s = s & "        "" et affich"" & ChrW(233) & "" une seule fois. L'utilisateur devra le changer "" & ChrW(224) & "" sa premi"" & ChrW(232) & ""re connexion.""" & vbCrLf
    s = s & "    FillRoles 0" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Changement de role d'un compte existant" & vbCrLf
    s = s & "Public Sub PrepareRole(ByVal pUserId As Long, ByVal pLogin As String, ByVal pFullName As String, ByVal pRoleId As Long)" & vbCrLf
    s = s & "    mblnRoleMode = True" & vbCrLf
    s = s & "    mlngUserId = pUserId" & vbCrLf
    s = s & "    Me.Caption = ""Kamerix - R"" & ChrW(244) & ""le du compte""" & vbCrLf
    s = s & "    lblTitle.Caption = ""Changer le r"" & ChrW(244) & ""le""" & vbCrLf
    s = s & "    lblInfo.Caption = ""Le nouveau r"" & ChrW(244) & ""le prend effet "" & ChrW(224) & "" la prochaine connexion de l'utilisateur.""" & vbCrLf
    s = s & "    txtLogin.Text = pLogin" & vbCrLf
    s = s & "    txtFullName.Text = pFullName" & vbCrLf
    s = s & "    txtLogin.Enabled = False" & vbCrLf
    s = s & "    txtFullName.Enabled = False" & vbCrLf
    s = s & "    FillRoles pRoleId" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_Initialize()" & vbCrLf
    s = s & "    lblMessage.Caption = """"" & vbCrLf
    s = s & "    cboRole.Style = 2              ' liste deroulante sans saisie libre" & vbCrLf
    s = s & "    cboRole.ColumnCount = 2        ' colonne 0 : RoleID (masquee), colonne 1 : nom" & vbCrLf
    s = s & "    cboRole.ColumnWidths = ""0 pt;150 pt""" & vbCrLf
    s = s & "    cboRole.BoundColumn = 1" & vbCrLf
    s = s & "    cboRole.TextColumn = 2" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub FillRoles(ByVal pSelectedRoleId As Long)" & vbCrLf
    s = s & "    Dim AdoRs As ADODB.Recordset" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    cboRole.Clear" & vbCrLf
    s = s & "    Set AdoRs = RoleList()" & vbCrLf
    s = s & "    Do While Not AdoRs.EOF" & vbCrLf
    s = s & "        cboRole.AddItem CStr(AdoRs.Fields(""RoleID"").Value)" & vbCrLf
    s = s & "        i = cboRole.ListCount - 1" & vbCrLf
    s = s & "        cboRole.List(i, 1) = CStr(AdoRs.Fields(""Nom"").Value)" & vbCrLf
    s = s & "        If CLng(AdoRs.Fields(""RoleID"").Value) = pSelectedRoleId Then cboRole.ListIndex = i" & vbCrLf
    s = s & "        AdoRs.MoveNext" & vbCrLf
    s = s & "    Loop" & vbCrLf
    s = s & "    AdoRs.Close" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdOK_Click()" & vbCrLf
    s = s & "    Dim lngRoleId As Long" & vbCrLf
    s = s & "    Dim strTemp As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    lblMessage.Caption = """"" & vbCrLf
    s = s & "    If cboRole.ListIndex < 0 Then" & vbCrLf
    s = s & "        lblMessage.Caption = ""Choisissez un r"" & ChrW(244) & ""le.""" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    lngRoleId = CLng(cboRole.List(cboRole.ListIndex, 0))" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo EH" & vbCrLf
    s = s & "    If mblnRoleMode Then" & vbCrLf
    s = s & "        ChangeAccountRole mlngUserId, lngRoleId" & vbCrLf
    s = s & "        mblnSucceeded = True" & vbCrLf
    s = s & "        Me.Hide" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        Me.MousePointer = 11" & vbCrLf
    s = s & "        CreateAccountWithTemporaryPassword txtLogin.Text, txtFullName.Text, lngRoleId, strTemp" & vbCrLf
    s = s & "        Me.MousePointer = 0" & vbCrLf
    s = s & "        mblnSucceeded = True" & vbCrLf
    s = s & "        Me.Hide" & vbCrLf
    s = s & "        ShowTemporaryPassword Trim$(txtLogin.Text), strTemp" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "EH:" & vbCrLf
    s = s & "    Me.MousePointer = 0" & vbCrLf
    s = s & "    lblMessage.Caption = Err.Description" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdCancel_Click()" & vbCrLf
    s = s & "    Me.Hide" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & vbCrLf
    s = s & "    If CloseMode = 0 Then" & vbCrLf
    s = s & "        Cancel = True" & vbCrLf
    s = s & "        Me.Hide" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    CodeFrmUserEdit = s
End Function

Private Function CodeFrmReference() As String
    Dim s As String

    s = s & "Option Explicit" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' KAMERIX - Referentiel : marques, modeles, villes, quartiers, commissariats (cree par" & vbCrLf
    s = s & "' ZmodFormBuilder : ne pas modifier a la main, modifier le generateur puis relancer BuildKamerixForms)" & vbCrLf
    s = s & "' La liste est chargee directement depuis la base (FillListControl : GetRows -> Column)." & vbCrLf
    s = s & "' Modifications : ecriture sur le module Voitures." & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Const KIND_BRANDS As Long = 0" & vbCrLf
    s = s & "Private Const KIND_MODELS As Long = 1" & vbCrLf
    s = s & "Private Const KIND_CITIES As Long = 2" & vbCrLf
    s = s & "Private Const KIND_DISTRICTS As Long = 3" & vbCrLf
    s = s & "Private Const KIND_STATIONS As Long = 4" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private mblnLoading As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Toujours Vrai (lu par OpenScreen)" & vbCrLf
    s = s & "Public Property Get Succeeded() As Boolean" & vbCrLf
    s = s & "    Succeeded = True" & vbCrLf
    s = s & "End Property" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_Initialize()" & vbCrLf
    s = s & "    mblnLoading = True" & vbCrLf
    s = s & "    cboKind.Style = 2" & vbCrLf
    s = s & "    cboKind.AddItem ""Marques""" & vbCrLf
    s = s & "    cboKind.AddItem ""Mod"" & ChrW(232) & ""les""" & vbCrLf
    s = s & "    cboKind.AddItem ""Villes""" & vbCrLf
    s = s & "    cboKind.AddItem ""Quartiers""" & vbCrLf
    s = s & "    cboKind.AddItem ""Commissariats""" & vbCrLf
    s = s & "    cboParent.Style = 2" & vbCrLf
    s = s & "    cboParent.ColumnCount = 2" & vbCrLf
    s = s & "    cboParent.ColumnWidths = ""0 pt;150 pt""" & vbCrLf
    s = s & "    cboParent.BoundColumn = 1" & vbCrLf
    s = s & "    cboParent.TextColumn = 2" & vbCrLf
    s = s & "    lblMessage.Caption = """"" & vbCrLf
    s = s & "    mblnLoading = False" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    cboKind.ListIndex = KIND_BRANDS          ' declenche cboKind_Change" & vbCrLf
    s = s & "    UpdateButtons" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' ---------- Choix de la liste ----------" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cboKind_Change()" & vbCrLf
    s = s & "    If mblnLoading Then Exit Sub" & vbCrLf
    s = s & "    lblMessage.Caption = """"" & vbCrLf
    s = s & "    txtName.Text = """"" & vbCrLf
    s = s & "    LoadParents" & vbCrLf
    s = s & "    LoadList" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cboParent_Change()" & vbCrLf
    s = s & "    If mblnLoading Then Exit Sub" & vbCrLf
    s = s & "    lblMessage.Caption = """"" & vbCrLf
    s = s & "    LoadList" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Liste des parents : marques (pour les modeles) ou villes (quartiers, commissariats)" & vbCrLf
    s = s & "Private Sub LoadParents()" & vbCrLf
    s = s & "    Dim AdoRs As ADODB.Recordset" & vbCrLf
    s = s & "    Dim strIdField As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    mblnLoading = True" & vbCrLf
    s = s & "    cboParent.Clear" & vbCrLf
    s = s & "    Select Case cboKind.ListIndex" & vbCrLf
    s = s & "        Case KIND_MODELS" & vbCrLf
    s = s & "            lblParent.Caption = ""Marque :""" & vbCrLf
    s = s & "            Set AdoRs = BrandList()" & vbCrLf
    s = s & "            strIdField = ""MarqueID""" & vbCrLf
    s = s & "        Case KIND_DISTRICTS, KIND_STATIONS" & vbCrLf
    s = s & "            lblParent.Caption = ""Ville :""" & vbCrLf
    s = s & "            Set AdoRs = PlaceList(PLACE_TYPE_CITY)" & vbCrLf
    s = s & "            strIdField = ""LieuID""" & vbCrLf
    s = s & "            If cboKind.ListIndex = KIND_STATIONS Then" & vbCrLf
    s = s & "                cboParent.AddItem ""0""" & vbCrLf
    s = s & "                cboParent.List(0, 1) = ""(toutes / aucune)""" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "    End Select" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    lblParent.Visible = Not (AdoRs Is Nothing)" & vbCrLf
    s = s & "    cboParent.Visible = Not (AdoRs Is Nothing)" & vbCrLf
    s = s & "    If Not AdoRs Is Nothing Then" & vbCrLf
    s = s & "        Do While Not AdoRs.EOF" & vbCrLf
    s = s & "            cboParent.AddItem CStr(AdoRs.Fields(strIdField).Value)" & vbCrLf
    s = s & "            cboParent.List(cboParent.ListCount - 1, 1) = CStr(AdoRs.Fields(""Nom"").Value)" & vbCrLf
    s = s & "            AdoRs.MoveNext" & vbCrLf
    s = s & "        Loop" & vbCrLf
    s = s & "        AdoRs.Close" & vbCrLf
    s = s & "        If cboParent.ListCount > 0 Then cboParent.ListIndex = 0" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    mblnLoading = False" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Recharge la liste affichee depuis la base ; en-tetes de colonnes dans lblCol1 et lblCol2" & vbCrLf
    s = s & "Private Sub LoadList()" & vbCrLf
    s = s & "    Dim lngParent As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    lngParent = ParentId()" & vbCrLf
    s = s & "    lstItems.Clear" & vbCrLf
    s = s & "    lblCol2.Visible = False" & vbCrLf
    s = s & "    lstItems.ColumnWidths = ""0 pt;250 pt""" & vbCrLf
    s = s & "    Select Case cboKind.ListIndex" & vbCrLf
    s = s & "        Case KIND_BRANDS" & vbCrLf
    s = s & "            lblCol1.Caption = ""Marque""" & vbCrLf
    s = s & "            FillListControl lstItems, BrandList(), Array(""MarqueID"", ""Nom"")" & vbCrLf
    s = s & "        Case KIND_MODELS" & vbCrLf
    s = s & "            lblCol1.Caption = ""Mod"" & ChrW(232) & ""le""" & vbCrLf
    s = s & "            If lngParent <> 0 Then FillListControl lstItems, ModelList(lngParent), Array(""ModeleID"", ""Nom"")" & vbCrLf
    s = s & "        Case KIND_CITIES" & vbCrLf
    s = s & "            lblCol1.Caption = ""Ville""" & vbCrLf
    s = s & "            FillListControl lstItems, PlaceList(PLACE_TYPE_CITY), Array(""LieuID"", ""Nom"")" & vbCrLf
    s = s & "        Case KIND_DISTRICTS" & vbCrLf
    s = s & "            lblCol1.Caption = ""Quartier""" & vbCrLf
    s = s & "            If lngParent <> 0 Then FillListControl lstItems, PlaceList(PLACE_TYPE_DISTRICT, lngParent), Array(""LieuID"", ""Nom"")" & vbCrLf
    s = s & "        Case KIND_STATIONS" & vbCrLf
    s = s & "            lblCol1.Caption = ""Commissariat""" & vbCrLf
    s = s & "            lblCol2.Caption = ""Ville""" & vbCrLf
    s = s & "            lblCol2.Visible = True" & vbCrLf
    s = s & "            lstItems.ColumnWidths = ""0 pt;160 pt;90 pt""" & vbCrLf
    s = s & "            FillListControl lstItems, PlaceList(PLACE_TYPE_POLICE_STATION, lngParent), Array(""LieuID"", ""Nom"", ""NomParent"")" & vbCrLf
    s = s & "    End Select" & vbCrLf
    s = s & "    UpdateButtons" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' ---------- Selection et boutons ----------" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub lstItems_Click()" & vbCrLf
    s = s & "    lblMessage.Caption = """"" & vbCrLf
    s = s & "    If lstItems.ListIndex >= 0 Then txtName.Text = CStr(lstItems.Column(1, lstItems.ListIndex))" & vbCrLf
    s = s & "    UpdateButtons" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UpdateButtons()" & vbCrLf
    s = s & "    Dim blnWrite As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    blnWrite = CanWrite(MOD_CARS)" & vbCrLf
    s = s & "    cmdAdd.Enabled = blnWrite" & vbCrLf
    s = s & "    cmdRename.Enabled = blnWrite And lstItems.ListIndex >= 0" & vbCrLf
    s = s & "    cmdDelete.Enabled = blnWrite And lstItems.ListIndex >= 0" & vbCrLf
    s = s & "    txtName.Enabled = blnWrite" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdAdd_Click()" & vbCrLf
    s = s & "    On Error GoTo EH" & vbCrLf
    s = s & "    RequireWrite MOD_CARS" & vbCrLf
    s = s & "    Select Case cboKind.ListIndex" & vbCrLf
    s = s & "        Case KIND_BRANDS:    AddBrand txtName.Text" & vbCrLf
    s = s & "        Case KIND_MODELS:    AddModel ParentId(), txtName.Text" & vbCrLf
    s = s & "        Case KIND_CITIES:    AddPlace txtName.Text, PLACE_TYPE_CITY" & vbCrLf
    s = s & "        Case KIND_DISTRICTS: AddPlace txtName.Text, PLACE_TYPE_DISTRICT, ParentId()" & vbCrLf
    s = s & "        Case KIND_STATIONS:  AddPlace txtName.Text, PLACE_TYPE_POLICE_STATION, ParentId()" & vbCrLf
    s = s & "    End Select" & vbCrLf
    s = s & "    lblMessage.Caption = ""Ajout"" & ChrW(233) & "" : "" & Trim$(txtName.Text)" & vbCrLf
    s = s & "    txtName.Text = """"" & vbCrLf
    s = s & "    LoadList" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "EH:" & vbCrLf
    s = s & "    lblMessage.Caption = Err.Description" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdRename_Click()" & vbCrLf
    s = s & "    Dim lngId As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo EH" & vbCrLf
    s = s & "    RequireWrite MOD_CARS" & vbCrLf
    s = s & "    lngId = SelectedId()" & vbCrLf
    s = s & "    Select Case cboKind.ListIndex" & vbCrLf
    s = s & "        Case KIND_BRANDS:  RenameBrand lngId, txtName.Text" & vbCrLf
    s = s & "        Case KIND_MODELS:  RenameModel lngId, txtName.Text" & vbCrLf
    s = s & "        Case Else:         RenamePlace lngId, txtName.Text" & vbCrLf
    s = s & "    End Select" & vbCrLf
    s = s & "    lblMessage.Caption = ""Renomm"" & ChrW(233) & "" : "" & Trim$(txtName.Text)" & vbCrLf
    s = s & "    LoadList" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "EH:" & vbCrLf
    s = s & "    lblMessage.Caption = Err.Description" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdDelete_Click()" & vbCrLf
    s = s & "    Dim lngId As Long" & vbCrLf
    s = s & "    Dim strName As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo EH" & vbCrLf
    s = s & "    RequireWrite MOD_CARS" & vbCrLf
    s = s & "    lngId = SelectedId()" & vbCrLf
    s = s & "    strName = CStr(lstItems.Column(1, lstItems.ListIndex))" & vbCrLf
    s = s & "    If MsgBox(""Supprimer "" & strName & "" ?"", vbQuestion + vbYesNo, ""Kamerix"") <> vbYes Then Exit Sub" & vbCrLf
    s = s & "    Select Case cboKind.ListIndex" & vbCrLf
    s = s & "        Case KIND_BRANDS:  DeleteBrand lngId" & vbCrLf
    s = s & "        Case KIND_MODELS:  DeleteModel lngId" & vbCrLf
    s = s & "        Case Else:         DeletePlace lngId" & vbCrLf
    s = s & "    End Select" & vbCrLf
    s = s & "    lblMessage.Caption = ""Supprim"" & ChrW(233) & "" : "" & strName" & vbCrLf
    s = s & "    txtName.Text = """"" & vbCrLf
    s = s & "    LoadList" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "EH:" & vbCrLf
    s = s & "    lblMessage.Caption = Err.Description" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdClose_Click()" & vbCrLf
    s = s & "    Me.Hide" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & vbCrLf
    s = s & "    If CloseMode = 0 Then" & vbCrLf
    s = s & "        Cancel = True" & vbCrLf
    s = s & "        Me.Hide" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' ---------- Outils ----------" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function SelectedId() As Long" & vbCrLf
    s = s & "    SelectedId = CLng(lstItems.Column(0, lstItems.ListIndex))" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Parent choisi (0 si aucun ou si la liste n'a pas de parent)" & vbCrLf
    s = s & "Private Function ParentId() As Long" & vbCrLf
    s = s & "    If cboParent.Visible And cboParent.ListIndex >= 0 Then ParentId = CLng(cboParent.List(cboParent.ListIndex, 0))" & vbCrLf
    s = s & "End Function" & vbCrLf
    CodeFrmReference = s
End Function

Private Function CodeFrmParties() As String
    Dim s As String

    s = s & "Option Explicit" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' KAMERIX - Tiers (acheteurs, vendeurs) et intermediaires (cree par ZmodFormBuilder : ne" & vbCrLf
    s = s & "' pas modifier a la main, modifier le generateur puis relancer BuildKamerixForms)" & vbCrLf
    s = s & "' Lecture : module Tiers (controle par OpenScreen). Ecriture : module Tiers." & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Const KIND_PARTIES As Long = 0" & vbCrLf
    s = s & "Private Const KIND_INTERMEDIARIES As Long = 1" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private mblnLoading As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Property Get Succeeded() As Boolean" & vbCrLf
    s = s & "    Succeeded = True" & vbCrLf
    s = s & "End Property" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_Initialize()" & vbCrLf
    s = s & "    mblnLoading = True" & vbCrLf
    s = s & "    cboKind.Style = 2" & vbCrLf
    s = s & "    cboKind.AddItem ""Acheteurs et vendeurs""" & vbCrLf
    s = s & "    cboKind.AddItem ""Interm"" & ChrW(233) & ""diaires""" & vbCrLf
    s = s & "    cboRole.Style = 2" & vbCrLf
    s = s & "    cboRole.ColumnCount = 2" & vbCrLf
    s = s & "    cboRole.ColumnWidths = ""0 pt;100 pt""" & vbCrLf
    s = s & "    cboRole.TextColumn = 2" & vbCrLf
    s = s & "    cboRole.AddItem PARTY_ROLE_ALL" & vbCrLf
    s = s & "    cboRole.List(0, 1) = ""Tous""" & vbCrLf
    s = s & "    cboRole.AddItem PARTY_ROLE_BUYER" & vbCrLf
    s = s & "    cboRole.List(1, 1) = ""Acheteurs""" & vbCrLf
    s = s & "    cboRole.AddItem PARTY_ROLE_SELLER" & vbCrLf
    s = s & "    cboRole.List(2, 1) = ""Vendeurs""" & vbCrLf
    s = s & "    cboRole.ListIndex = 0" & vbCrLf
    s = s & "    lblMessage.Caption = """"" & vbCrLf
    s = s & "    mblnLoading = False" & vbCrLf
    s = s & "    cboKind.ListIndex = KIND_PARTIES" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cboKind_Change()" & vbCrLf
    s = s & "    If mblnLoading Then Exit Sub" & vbCrLf
    s = s & "    lblRole.Visible = (cboKind.ListIndex = KIND_PARTIES)" & vbCrLf
    s = s & "    cboRole.Visible = lblRole.Visible" & vbCrLf
    s = s & "    LoadList" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cboRole_Change()" & vbCrLf
    s = s & "    If Not mblnLoading Then LoadList" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub txtSearch_Change()" & vbCrLf
    s = s & "    If Not mblnLoading Then LoadList" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Recharge la liste ; pSelectId = ligne a reselectionner (0 = aucune)" & vbCrLf
    s = s & "Private Sub LoadList(Optional ByVal pSelectId As Long = 0)" & vbCrLf
    s = s & "    On Error GoTo EH" & vbCrLf
    s = s & "    lblMessage.Caption = """"" & vbCrLf
    s = s & "    If cboKind.ListIndex = KIND_PARTIES Then" & vbCrLf
    s = s & "        lblCol1.Caption = ""Nom"": lblCol2.Caption = ""T"" & ChrW(233) & ""l"" & ChrW(233) & ""phone""" & vbCrLf
    s = s & "        lblCol3.Caption = ""Pi"" & ChrW(232) & ""ce"": lblCol4.Caption = ""R"" & ChrW(244) & ""le""" & vbCrLf
    s = s & "        FillListControl lstItems, PartyList(txtSearch.Text, CStr(cboRole.List(cboRole.ListIndex, 0))), _" & vbCrLf
    s = s & "                        Array(""TiersID"", ""Nom"", ""Telephone"", ""Piece"", ""Roles"")" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        lblCol1.Caption = ""Nom"": lblCol2.Caption = ""T"" & ChrW(233) & ""l"" & ChrW(233) & ""phone""" & vbCrLf
    s = s & "        lblCol3.Caption = ""Ville"": lblCol4.Caption = """"" & vbCrLf
    s = s & "        FillListControl lstItems, IntermediaryList(txtSearch.Text), Array(""IntermediaireID"", ""Nom"", ""Telephone"", ""NomVille"")" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    lstItems.ColumnWidths = ""0 pt;160 pt;95 pt;110 pt;105 pt""" & vbCrLf
    s = s & "    If pSelectId <> 0 Then SelectById lstItems, pSelectId" & vbCrLf
    s = s & "    UpdateButtons" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "EH:" & vbCrLf
    s = s & "    lblMessage.Caption = Err.Description" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub lstItems_Change()" & vbCrLf
    s = s & "    UpdateButtons" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub lstItems_DblClick(ByVal Cancel As MSForms.ReturnBoolean)" & vbCrLf
    s = s & "    If cmdEdit.Enabled Then cmdEdit_Click" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UpdateButtons()" & vbCrLf
    s = s & "    cmdNew.Enabled = CanWrite(MOD_PARTIES)" & vbCrLf
    s = s & "    cmdEdit.Enabled = (lstItems.ListIndex >= 0)" & vbCrLf
    s = s & "    If CanWrite(MOD_PARTIES) Then cmdEdit.Caption = ""Modifier..."" Else cmdEdit.Caption = ""Consulter...""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdNew_Click()" & vbCrLf
    s = s & "    OpenEditForm 0" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdEdit_Click()" & vbCrLf
    s = s & "    OpenEditForm SelectedIdOf(lstItems)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Ouvre la fiche (pId = 0 : nouvelle fiche)" & vbCrLf
    s = s & "Private Sub OpenEditForm(ByVal pId As Long)" & vbCrLf
    s = s & "    Dim frm As Object" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo EH" & vbCrLf
    s = s & "    If cboKind.ListIndex = KIND_PARTIES Then" & vbCrLf
    s = s & "        Set frm = VBA.UserForms.Add(""frmPartyEdit"")" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        Set frm = VBA.UserForms.Add(""frmIntermediaryEdit"")" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    frm.Prepare pId" & vbCrLf
    s = s & "    frm.Show" & vbCrLf
    s = s & "    If frm.Succeeded Then LoadList frm.SavedId" & vbCrLf
    s = s & "    Unload frm" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "EH:" & vbCrLf
    s = s & "    lblMessage.Caption = Err.Description" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdClose_Click()" & vbCrLf
    s = s & "    Me.Hide" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & vbCrLf
    s = s & "    If CloseMode = 0 Then" & vbCrLf
    s = s & "        Cancel = True" & vbCrLf
    s = s & "        Me.Hide" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    CodeFrmParties = s
End Function

Private Function CodeFrmPartyEdit() As String
    Dim s As String

    s = s & "Option Explicit" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' KAMERIX - Fiche d'un tiers (acheteur et/ou vendeur) (cree par ZmodFormBuilder : ne pas" & vbCrLf
    s = s & "' modifier a la main, modifier le generateur puis relancer BuildKamerixForms)" & vbCrLf
    s = s & "' Appeler Prepare (0 = nouveau tiers) avant Show." & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private mParty As Party" & vbCrLf
    s = s & "Private mblnSucceeded As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Property Get Succeeded() As Boolean" & vbCrLf
    s = s & "    Succeeded = mblnSucceeded" & vbCrLf
    s = s & "End Property" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Identifiant du tiers enregistre" & vbCrLf
    s = s & "Public Property Get SavedId() As Long" & vbCrLf
    s = s & "    If Not mParty Is Nothing Then SavedId = mParty.Id" & vbCrLf
    s = s & "End Property" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub Prepare(ByVal pPartyId As Long)" & vbCrLf
    s = s & "    lblMessage.Caption = """"" & vbCrLf
    s = s & "    If pPartyId = 0 Then" & vbCrLf
    s = s & "        Set mParty = New Party" & vbCrLf
    s = s & "        lblTitle.Caption = ""Nouveau tiers""" & vbCrLf
    s = s & "        chkBuyer.Value = True" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        Set mParty = LoadPartyChecked(pPartyId)" & vbCrLf
    s = s & "        lblTitle.Caption = ""Fiche du tiers""" & vbCrLf
    s = s & "        txtName.Text = mParty.Name" & vbCrLf
    s = s & "        txtPhone.Text = mParty.Phone" & vbCrLf
    s = s & "        txtIdType.Text = mParty.IdType" & vbCrLf
    s = s & "        txtIdNumber.Text = mParty.IdNumber" & vbCrLf
    s = s & "        txtAddress.Text = mParty.Address" & vbCrLf
    s = s & "        chkBuyer.Value = mParty.IsBuyer" & vbCrLf
    s = s & "        chkSeller.Value = mParty.IsSeller" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    cmdOK.Enabled = CanWrite(MOD_PARTIES)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdOK_Click()" & vbCrLf
    s = s & "    On Error GoTo EH" & vbCrLf
    s = s & "    lblMessage.Caption = """"" & vbCrLf
    s = s & "    mParty.Name = txtName.Text" & vbCrLf
    s = s & "    mParty.Phone = txtPhone.Text" & vbCrLf
    s = s & "    mParty.IdType = txtIdType.Text" & vbCrLf
    s = s & "    mParty.IdNumber = txtIdNumber.Text" & vbCrLf
    s = s & "    mParty.Address = txtAddress.Text" & vbCrLf
    s = s & "    mParty.IsBuyer = chkBuyer.Value" & vbCrLf
    s = s & "    mParty.IsSeller = chkSeller.Value" & vbCrLf
    s = s & "    SavePartyChecked mParty" & vbCrLf
    s = s & "    mblnSucceeded = True" & vbCrLf
    s = s & "    Me.Hide" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "EH:" & vbCrLf
    s = s & "    lblMessage.Caption = Err.Description" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdCancel_Click()" & vbCrLf
    s = s & "    Me.Hide" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & vbCrLf
    s = s & "    If CloseMode = 0 Then" & vbCrLf
    s = s & "        Cancel = True" & vbCrLf
    s = s & "        Me.Hide" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    CodeFrmPartyEdit = s
End Function

Private Function CodeFrmIntermediaryEdit() As String
    Dim s As String

    s = s & "Option Explicit" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' KAMERIX - Fiche d'un intermediaire (cree par ZmodFormBuilder : ne pas modifier a la" & vbCrLf
    s = s & "' main, modifier le generateur puis relancer BuildKamerixForms)" & vbCrLf
    s = s & "' Appeler Prepare (0 = nouvel intermediaire) avant Show." & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private mIntermediary As Intermediary" & vbCrLf
    s = s & "Private mblnSucceeded As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Property Get Succeeded() As Boolean" & vbCrLf
    s = s & "    Succeeded = mblnSucceeded" & vbCrLf
    s = s & "End Property" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Property Get SavedId() As Long" & vbCrLf
    s = s & "    If Not mIntermediary Is Nothing Then SavedId = mIntermediary.Id" & vbCrLf
    s = s & "End Property" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_Initialize()" & vbCrLf
    s = s & "    cboCity.Style = 2" & vbCrLf
    s = s & "    cboCity.ColumnWidths = ""0 pt;150 pt""" & vbCrLf
    s = s & "    cboCity.TextColumn = 2" & vbCrLf
    s = s & "    FillListControl cboCity, PlaceList(PLACE_TYPE_CITY), Array(""LieuID"", ""Nom"")" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub Prepare(ByVal pIntermediaryId As Long)" & vbCrLf
    s = s & "    lblMessage.Caption = """"" & vbCrLf
    s = s & "    If pIntermediaryId = 0 Then" & vbCrLf
    s = s & "        Set mIntermediary = New Intermediary" & vbCrLf
    s = s & "        lblTitle.Caption = ""Nouvel interm"" & ChrW(233) & ""diaire""" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        Set mIntermediary = LoadIntermediaryChecked(pIntermediaryId)" & vbCrLf
    s = s & "        lblTitle.Caption = ""Fiche de l'interm"" & ChrW(233) & ""diaire""" & vbCrLf
    s = s & "        txtName.Text = mIntermediary.Name" & vbCrLf
    s = s & "        txtPhone.Text = mIntermediary.Phone" & vbCrLf
    s = s & "        SelectById cboCity, mIntermediary.CityId" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If cboCity.ListCount = 0 Then lblMessage.Caption = ""Aucune ville : ajoutez-en dans le r"" & ChrW(233) & ""f"" & ChrW(233) & ""rentiel.""" & vbCrLf
    s = s & "    cmdOK.Enabled = CanWrite(MOD_PARTIES)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdOK_Click()" & vbCrLf
    s = s & "    On Error GoTo EH" & vbCrLf
    s = s & "    lblMessage.Caption = """"" & vbCrLf
    s = s & "    mIntermediary.Name = txtName.Text" & vbCrLf
    s = s & "    mIntermediary.Phone = txtPhone.Text" & vbCrLf
    s = s & "    mIntermediary.CityId = SelectedIdOf(cboCity)" & vbCrLf
    s = s & "    SaveIntermediaryChecked mIntermediary" & vbCrLf
    s = s & "    mblnSucceeded = True" & vbCrLf
    s = s & "    Me.Hide" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "EH:" & vbCrLf
    s = s & "    lblMessage.Caption = Err.Description" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdCancel_Click()" & vbCrLf
    s = s & "    Me.Hide" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & vbCrLf
    s = s & "    If CloseMode = 0 Then" & vbCrLf
    s = s & "        Cancel = True" & vbCrLf
    s = s & "        Me.Hide" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    CodeFrmIntermediaryEdit = s
End Function

Private Function CodeFrmCars() As String
    Dim s As String

    s = s & "Option Explicit" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' KAMERIX - Voitures : liste filtree par statut, nouvel achat, fiche, retrait du stock" & vbCrLf
    s = s & "' (cree par ZmodFormBuilder : ne pas modifier a la main, modifier le generateur puis" & vbCrLf
    s = s & "' relancer BuildKamerixForms). Lecture : module Voitures. Ecriture : module Voitures." & vbCrLf
    s = s & "' Colonnes de lstCars : 0 VoitureID (masquee), 1 modele, 2 annee, 3 chassis," & vbCrLf
    s = s & "' 4 immatriculation, 5 statut, 6 prix d'achat, 7 code du statut (masque)." & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Const COL_STATUS_CODE As Long = 7" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private mblnLoading As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Property Get Succeeded() As Boolean" & vbCrLf
    s = s & "    Succeeded = True" & vbCrLf
    s = s & "End Property" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_Initialize()" & vbCrLf
    s = s & "    Dim vCodes As Variant" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    mblnLoading = True" & vbCrLf
    s = s & "    cboStatus.Style = 2" & vbCrLf
    s = s & "    cboStatus.ColumnCount = 2" & vbCrLf
    s = s & "    cboStatus.ColumnWidths = ""0 pt;110 pt""" & vbCrLf
    s = s & "    cboStatus.TextColumn = 2" & vbCrLf
    s = s & "    vCodes = Array("""", CAR_STATUS_IN_STOCK, CAR_STATUS_RESERVED, CAR_STATUS_SALE_IN_PROGRESS, CAR_STATUS_SOLD, CAR_STATUS_WITHDRAWN)" & vbCrLf
    s = s & "    For i = LBound(vCodes) To UBound(vCodes)" & vbCrLf
    s = s & "        cboStatus.AddItem CStr(vCodes(i))" & vbCrLf
    s = s & "        If i = 0 Then cboStatus.List(i, 1) = ""Toutes"" Else cboStatus.List(i, 1) = CarStatusLabel(CStr(vCodes(i)))" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "    lblMessage.Caption = """"" & vbCrLf
    s = s & "    mblnLoading = False" & vbCrLf
    s = s & "    cboStatus.ListIndex = 1          ' en stock" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cboStatus_Change()" & vbCrLf
    s = s & "    If Not mblnLoading Then LoadList" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub LoadList(Optional ByVal pSelectId As Long = 0)" & vbCrLf
    s = s & "    Dim lngCount As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo EH" & vbCrLf
    s = s & "    lngCount = FillListControl(lstCars, CarList(CStr(cboStatus.List(cboStatus.ListIndex, 0))), _" & vbCrLf
    s = s & "        Array(""VoitureID"", ""Modele"", ""Annee"", ""Chassis"", ""Immatriculation"", ""LibelleStatut"", ""PrixAchat"", ""Statut""), _" & vbCrLf
    s = s & "        Array("""", """", """", """", """", """", ""#,##0"", """"))" & vbCrLf
    s = s & "    lstCars.ColumnWidths = ""0 pt;125 pt;35 pt;115 pt;85 pt;80 pt;70 pt;0 pt""" & vbCrLf
    s = s & "    lblCount.Caption = lngCount & "" voiture(s)""" & vbCrLf
    s = s & "    If pSelectId <> 0 Then SelectById lstCars, pSelectId" & vbCrLf
    s = s & "    UpdateButtons" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "EH:" & vbCrLf
    s = s & "    lblMessage.Caption = Err.Description" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub lstCars_Change()" & vbCrLf
    s = s & "    lblMessage.Caption = """"" & vbCrLf
    s = s & "    UpdateButtons" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub lstCars_DblClick(ByVal Cancel As MSForms.ReturnBoolean)" & vbCrLf
    s = s & "    If cmdEdit.Enabled Then cmdEdit_Click" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UpdateButtons()" & vbCrLf
    s = s & "    Dim blnWrite As Boolean" & vbCrLf
    s = s & "    Dim blnSelected As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    blnWrite = CanWrite(MOD_CARS)" & vbCrLf
    s = s & "    blnSelected = (lstCars.ListIndex >= 0)" & vbCrLf
    s = s & "    cmdPurchase.Enabled = blnWrite" & vbCrLf
    s = s & "    cmdEdit.Enabled = blnSelected" & vbCrLf
    s = s & "    If blnWrite Then cmdEdit.Caption = ""Modifier..."" Else cmdEdit.Caption = ""Consulter...""" & vbCrLf
    s = s & "    cmdWithdraw.Enabled = False" & vbCrLf
    s = s & "    If blnWrite And blnSelected Then" & vbCrLf
    s = s & "        cmdWithdraw.Enabled = (CStr(lstCars.List(lstCars.ListIndex, COL_STATUS_CODE)) = CAR_STATUS_IN_STOCK)" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdPurchase_Click()" & vbCrLf
    s = s & "    OpenCarForm 0" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdEdit_Click()" & vbCrLf
    s = s & "    OpenCarForm SelectedIdOf(lstCars)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Fiche voiture : pCarId = 0 pour un nouvel achat" & vbCrLf
    s = s & "Private Sub OpenCarForm(ByVal pCarId As Long)" & vbCrLf
    s = s & "    Dim frm As Object" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo EH" & vbCrLf
    s = s & "    Set frm = VBA.UserForms.Add(""frmCarEdit"")" & vbCrLf
    s = s & "    frm.Prepare pCarId" & vbCrLf
    s = s & "    frm.Show" & vbCrLf
    s = s & "    If frm.Succeeded Then LoadList frm.SavedId" & vbCrLf
    s = s & "    Unload frm" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "EH:" & vbCrLf
    s = s & "    lblMessage.Caption = Err.Description" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdWithdraw_Click()" & vbCrLf
    s = s & "    Dim lngId As Long" & vbCrLf
    s = s & "    Dim strReason As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo EH" & vbCrLf
    s = s & "    lngId = SelectedIdOf(lstCars)" & vbCrLf
    s = s & "    strReason = Trim$(InputBox(""Motif du retrait du stock (100 caract"" & ChrW(232) & ""res au plus) :"", ""Kamerix - Retrait""))" & vbCrLf
    s = s & "    If Len(strReason) = 0 Then Exit Sub" & vbCrLf
    s = s & "    If MsgBox(""Retirer cette voiture du stock ?"" & vbCrLf & ""Ce statut est d"" & ChrW(233) & ""finitif."", _" & vbCrLf
    s = s & "              vbQuestion + vbYesNo, ""Kamerix"") <> vbYes Then Exit Sub" & vbCrLf
    s = s & "    WithdrawCar lngId, strReason" & vbCrLf
    s = s & "    LoadList" & vbCrLf
    s = s & "    lblMessage.Caption = ""Voiture retir"" & ChrW(233) & ""e du stock.""" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "EH:" & vbCrLf
    s = s & "    lblMessage.Caption = Err.Description" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdClose_Click()" & vbCrLf
    s = s & "    Me.Hide" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & vbCrLf
    s = s & "    If CloseMode = 0 Then" & vbCrLf
    s = s & "        Cancel = True" & vbCrLf
    s = s & "        Me.Hide" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    CodeFrmCars = s
End Function

Private Function CodeFrmCarEdit() As String
    Dim s As String

    s = s & "Option Explicit" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' KAMERIX - Fiche voiture : nouvel achat (voiture + achat) ou modification des donnees" & vbCrLf
    s = s & "' (cree par ZmodFormBuilder : ne pas modifier a la main, modifier le generateur puis" & vbCrLf
    s = s & "' relancer BuildKamerixForms). Appeler Prepare (0 = nouvel achat) avant Show." & vbCrLf
    s = s & "' Une voiture existante : seules ses donnees descriptives se modifient ; l'achat est" & vbCrLf
    s = s & "' affiche en consultation." & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private mCar As Car" & vbCrLf
    s = s & "Private mblnNew As Boolean" & vbCrLf
    s = s & "Private mblnSucceeded As Boolean" & vbCrLf
    s = s & "Private mblnLoading As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Property Get Succeeded() As Boolean" & vbCrLf
    s = s & "    Succeeded = mblnSucceeded" & vbCrLf
    s = s & "End Property" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Property Get SavedId() As Long" & vbCrLf
    s = s & "    If Not mCar Is Nothing Then SavedId = mCar.Id" & vbCrLf
    s = s & "End Property" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_Initialize()" & vbCrLf
    s = s & "    mblnLoading = True" & vbCrLf
    s = s & "    SetupCombo cboBrand" & vbCrLf
    s = s & "    SetupCombo cboModel" & vbCrLf
    s = s & "    SetupCombo cboSeller" & vbCrLf
    s = s & "    SetupCombo cboIntermediary" & vbCrLf
    s = s & "    SetupCombo cboPlace" & vbCrLf
    s = s & "    FillListControl cboBrand, BrandList(), Array(""MarqueID"", ""Nom"")" & vbCrLf
    s = s & "    FillListControl cboSeller, PartyList("""", PARTY_ROLE_SELLER), Array(""TiersID"", ""Nom"")" & vbCrLf
    s = s & "    FillListControl cboIntermediary, IntermediaryList(), Array(""IntermediaireID"", ""Nom"")" & vbCrLf
    s = s & "    AddNoneItem cboIntermediary" & vbCrLf
    s = s & "    FillListControl cboPlace, PlaceList(PLACE_TYPE_CITY), Array(""LieuID"", ""Nom"")" & vbCrLf
    s = s & "    AddNoneItem cboPlace" & vbCrLf
    s = s & "    lblMessage.Caption = """"" & vbCrLf
    s = s & "    mblnLoading = False" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub SetupCombo(ByVal pCombo As Object)" & vbCrLf
    s = s & "    pCombo.Style = 2" & vbCrLf
    s = s & "    pCombo.ColumnCount = 2" & vbCrLf
    s = s & "    pCombo.ColumnWidths = ""0 pt;160 pt""" & vbCrLf
    s = s & "    pCombo.TextColumn = 2" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub Prepare(ByVal pCarId As Long)" & vbCrLf
    s = s & "    Dim objPurchase As Movement" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    mblnNew = (pCarId = 0)" & vbCrLf
    s = s & "    If mblnNew Then" & vbCrLf
    s = s & "        Set mCar = New Car" & vbCrLf
    s = s & "        Me.Caption = ""Kamerix - Nouvel achat""" & vbCrLf
    s = s & "        lblTitle.Caption = ""Nouvel achat""" & vbCrLf
    s = s & "        cboIntermediary.ListIndex = 0" & vbCrLf
    s = s & "        cboPlace.ListIndex = 0" & vbCrLf
    s = s & "        txtDate.Text = Format$(Date, ""dd/mm/yyyy"")" & vbCrLf
    s = s & "        txtCommission.Text = ""0""" & vbCrLf
    s = s & "        cmdOK.Enabled = CanWrite(MOD_CARS)" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set mCar = LoadCarChecked(pCarId)" & vbCrLf
    s = s & "    Me.Caption = ""Kamerix - Fiche voiture""" & vbCrLf
    s = s & "    lblTitle.Caption = ""Fiche voiture : "" & CarStatusLabel(mCar.Status)" & vbCrLf
    s = s & "    mblnLoading = True" & vbCrLf
    s = s & "    SelectById cboBrand, BrandOfModel(mCar.ModelId)" & vbCrLf
    s = s & "    LoadModels" & vbCrLf
    s = s & "    SelectById cboModel, mCar.ModelId" & vbCrLf
    s = s & "    mblnLoading = False" & vbCrLf
    s = s & "    If Not IsNull(mCar.ModelYear) Then txtYear.Text = CStr(mCar.ModelYear)" & vbCrLf
    s = s & "    txtColor.Text = mCar.Color" & vbCrLf
    s = s & "    txtChassis.Text = mCar.Chassis" & vbCrLf
    s = s & "    txtPlate.Text = mCar.PlateNumber" & vbCrLf
    s = s & "    If Not IsNull(mCar.Mileage) Then txtMileage.Text = Format$(mCar.Mileage, ""#,##0"")" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' Achat : consultation seule" & vbCrLf
    s = s & "    Set objPurchase = PurchaseOfCar(pCarId)" & vbCrLf
    s = s & "    If Not objPurchase Is Nothing Then" & vbCrLf
    s = s & "        SelectById cboSeller, objPurchase.PartyId" & vbCrLf
    s = s & "        txtPrice.Text = Format$(objPurchase.Price, ""#,##0"")" & vbCrLf
    s = s & "        SelectById cboIntermediary, objPurchase.IntermediaryId" & vbCrLf
    s = s & "        txtCommission.Text = Format$(objPurchase.Commission, ""#,##0"")" & vbCrLf
    s = s & "        SelectById cboPlace, objPurchase.PlaceId" & vbCrLf
    s = s & "        txtDate.Text = Format$(objPurchase.MovementDate, ""dd/mm/yyyy"")" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        lblMessage.Caption = ""Achat non enregistr"" & ChrW(233) & "" pour cette voiture.""" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    LockPurchase" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If mCar.Status = CAR_STATUS_SOLD Or mCar.Status = CAR_STATUS_WITHDRAWN Then" & vbCrLf
    s = s & "        cmdOK.Enabled = False" & vbCrLf
    s = s & "        lblMessage.Caption = ""Voiture vendue ou retir"" & ChrW(233) & ""e : consultation seule.""" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        cmdOK.Enabled = CanWrite(MOD_CARS)" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub LockPurchase()" & vbCrLf
    s = s & "    cboSeller.Enabled = False" & vbCrLf
    s = s & "    txtPrice.Enabled = False" & vbCrLf
    s = s & "    cboIntermediary.Enabled = False" & vbCrLf
    s = s & "    txtCommission.Enabled = False" & vbCrLf
    s = s & "    cboPlace.Enabled = False" & vbCrLf
    s = s & "    txtDate.Enabled = False" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cboBrand_Change()" & vbCrLf
    s = s & "    If Not mblnLoading Then LoadModels" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub LoadModels()" & vbCrLf
    s = s & "    cboModel.Clear" & vbCrLf
    s = s & "    If SelectedIdOf(cboBrand) <> 0 Then" & vbCrLf
    s = s & "        FillListControl cboModel, ModelList(SelectedIdOf(cboBrand)), Array(""ModeleID"", ""Nom"")" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdOK_Click()" & vbCrLf
    s = s & "    Dim curPrice As Currency" & vbCrLf
    s = s & "    Dim curCommission As Currency" & vbCrLf
    s = s & "    Dim curMileage As Currency" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo EH" & vbCrLf
    s = s & "    lblMessage.Caption = """"" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' Donnees de la voiture (la classe Car valide et signale toutes les erreurs)" & vbCrLf
    s = s & "    mCar.ModelId = SelectedIdOf(cboModel)" & vbCrLf
    s = s & "    mCar.ModelYear = Trim$(txtYear.Text)" & vbCrLf
    s = s & "    mCar.Color = txtColor.Text" & vbCrLf
    s = s & "    mCar.Chassis = txtChassis.Text" & vbCrLf
    s = s & "    mCar.PlateNumber = txtPlate.Text" & vbCrLf
    s = s & "    If Not TryParseAmount(txtMileage.Text, curMileage) Then" & vbCrLf
    s = s & "        lblMessage.Caption = ""Le kilom"" & ChrW(233) & ""trage doit "" & ChrW(234) & ""tre un nombre.""" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If Len(Trim$(txtMileage.Text)) = 0 Then mCar.Mileage = Null Else mCar.Mileage = curMileage" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If Not mblnNew Then" & vbCrLf
    s = s & "        SaveCarDetails mCar" & vbCrLf
    s = s & "        mblnSucceeded = True" & vbCrLf
    s = s & "        Me.Hide" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    ' Nouvel achat" & vbCrLf
    s = s & "    If SelectedIdOf(cboSeller) = 0 Then" & vbCrLf
    s = s & "        lblMessage.Caption = ""Choisissez le vendeur.""" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If Not TryParseAmount(txtPrice.Text, curPrice) Or Not TryParseAmount(txtCommission.Text, curCommission) Then" & vbCrLf
    s = s & "        lblMessage.Caption = ""Le prix et la commission doivent "" & ChrW(234) & ""tre des montants.""" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If Not IsDate(txtDate.Text) Then" & vbCrLf
    s = s & "        lblMessage.Caption = ""La date de l'achat est invalide (jj/mm/aaaa).""" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    RecordPurchase mCar, SelectedIdOf(cboSeller), curPrice, curCommission, SelectedIdOf(cboIntermediary), _" & vbCrLf
    s = s & "                   SelectedIdOf(cboPlace), CDate(txtDate.Text)" & vbCrLf
    s = s & "    mblnSucceeded = True" & vbCrLf
    s = s & "    Me.Hide" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "EH:" & vbCrLf
    s = s & "    lblMessage.Caption = Err.Description" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdCancel_Click()" & vbCrLf
    s = s & "    Me.Hide" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & vbCrLf
    s = s & "    If CloseMode = 0 Then" & vbCrLf
    s = s & "        Cancel = True" & vbCrLf
    s = s & "        Me.Hide" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    CodeFrmCarEdit = s
End Function

Private Function CodeFrmSales() As String
    Dim s As String

    s = s & "Option Explicit" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' KAMERIX - Ventes comptant : liste filtree par statut, nouvelle vente, paiement, CV," & vbCrLf
    s = s & "' annulation (cree par ZmodFormBuilder : ne pas modifier a la main, modifier le" & vbCrLf
    s = s & "' generateur puis relancer BuildKamerixForms). Lecture : module Ventes." & vbCrLf
    s = s & "' Colonnes de lstSales : voir les constantes SALE_COL_* de modKamerixSales." & vbCrLf
    s = s & "' Documents... : facture, recus et remboursements de la vente (frmSaleDocuments, lot 5.5)." & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private mblnLoading As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Property Get Succeeded() As Boolean" & vbCrLf
    s = s & "    Succeeded = True" & vbCrLf
    s = s & "End Property" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_Initialize()" & vbCrLf
    s = s & "    Dim vCodes As Variant" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    mblnLoading = True" & vbCrLf
    s = s & "    cboStatus.Style = 2" & vbCrLf
    s = s & "    cboStatus.ColumnCount = 2" & vbCrLf
    s = s & "    cboStatus.ColumnWidths = ""0 pt;100 pt""" & vbCrLf
    s = s & "    cboStatus.TextColumn = 2" & vbCrLf
    s = s & "    vCodes = Array(MOVEMENT_STATUS_IN_PROGRESS, MOVEMENT_STATUS_SETTLED, MOVEMENT_STATUS_CANCELLED, """")" & vbCrLf
    s = s & "    For i = LBound(vCodes) To UBound(vCodes)" & vbCrLf
    s = s & "        cboStatus.AddItem CStr(vCodes(i))" & vbCrLf
    s = s & "        If Len(vCodes(i)) = 0 Then cboStatus.List(i, 1) = ""Toutes"" Else cboStatus.List(i, 1) = SaleStatusLabel(CStr(vCodes(i)))" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "    lblMessage.Caption = """"" & vbCrLf
    s = s & "    mblnLoading = False" & vbCrLf
    s = s & "    cboStatus.ListIndex = 0          ' en cours" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cboStatus_Change()" & vbCrLf
    s = s & "    If Not mblnLoading Then LoadList" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub LoadList(Optional ByVal pSelectId As Long = 0)" & vbCrLf
    s = s & "    Dim vData As Variant" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo EH" & vbCrLf
    s = s & "    lstSales.Clear" & vbCrLf
    s = s & "    lstSales.ColumnCount = 11" & vbCrLf
    s = s & "    lstSales.ColumnWidths = ""0 pt;55 pt;110 pt;150 pt;65 pt;65 pt;65 pt;55 pt;55 pt;0 pt;0 pt""" & vbCrLf
    s = s & "    vData = SaleRows(CStr(cboStatus.List(cboStatus.ListIndex, 0)))" & vbCrLf
    s = s & "    If IsEmpty(vData) Then" & vbCrLf
    s = s & "        lblCount.Caption = ""0 vente""" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        lstSales.Column = vData" & vbCrLf
    s = s & "        lblCount.Caption = (UBound(vData, 2) + 1) & "" vente(s)""" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If pSelectId <> 0 Then SelectById lstSales, pSelectId" & vbCrLf
    s = s & "    UpdateButtons" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "EH:" & vbCrLf
    s = s & "    lblMessage.Caption = Err.Description" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub lstSales_Change()" & vbCrLf
    s = s & "    UpdateButtons" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UpdateButtons()" & vbCrLf
    s = s & "    Dim strStatus As String" & vbCrLf
    s = s & "    Dim strCv As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    cmdNew.Enabled = CanWrite(MOD_SALES)" & vbCrLf
    s = s & "    cmdPay.Enabled = False" & vbCrLf
    s = s & "    cmdCv.Enabled = False" & vbCrLf
    s = s & "    cmdCancel.Enabled = False" & vbCrLf
    s = s & "    cmdDocs.Enabled = False" & vbCrLf
    s = s & "    If lstSales.ListIndex < 0 Then Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    strStatus = CStr(lstSales.List(lstSales.ListIndex, SALE_COL_STATUS_CODE))" & vbCrLf
    s = s & "    strCv = CStr(lstSales.List(lstSales.ListIndex, SALE_COL_CV_CODE))" & vbCrLf
    s = s & "    cmdPay.Enabled = CanWrite(MOD_PAYMENTS) And strStatus = MOVEMENT_STATUS_IN_PROGRESS" & vbCrLf
    s = s & "    cmdCv.Enabled = (strStatus <> MOVEMENT_STATUS_CANCELLED)" & vbCrLf
    s = s & "    cmdCancel.Enabled = HasRight(RIGHT_CANCEL_SALE) And strStatus <> MOVEMENT_STATUS_CANCELLED And strCv <> CV_STATUS_SIGNED" & vbCrLf
    s = s & "    cmdDocs.Enabled = CanReadDocument(DOC_INVOICE) Or CanReadDocument(DOC_RECEIPT)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdNew_Click()" & vbCrLf
    s = s & "    Dim frm As Object" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo EH" & vbCrLf
    s = s & "    Set frm = VBA.UserForms.Add(""frmSaleStart"")" & vbCrLf
    s = s & "    frm.Show" & vbCrLf
    s = s & "    If frm.Succeeded Then" & vbCrLf
    s = s & "        mblnLoading = True" & vbCrLf
    s = s & "        cboStatus.ListIndex = 0" & vbCrLf
    s = s & "        mblnLoading = False" & vbCrLf
    s = s & "        LoadList frm.SavedId" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Unload frm" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "EH:" & vbCrLf
    s = s & "    lblMessage.Caption = Err.Description" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdPay_Click()" & vbCrLf
    s = s & "    OpenSaleForm ""frmSalePayment""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdCv_Click()" & vbCrLf
    s = s & "    OpenSaleForm ""frmSaleCv""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdCancel_Click()" & vbCrLf
    s = s & "    OpenSaleForm ""frmSaleCancel""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdDocs_Click()" & vbCrLf
    s = s & "    OpenSaleForm ""frmSaleDocuments""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Ouvre un formulaire qui agit sur la vente selectionnee, puis recharge la liste" & vbCrLf
    s = s & "Private Sub OpenSaleForm(ByVal pFormName As String)" & vbCrLf
    s = s & "    Dim frm As Object" & vbCrLf
    s = s & "    Dim lngId As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo EH" & vbCrLf
    s = s & "    lngId = SelectedIdOf(lstSales)" & vbCrLf
    s = s & "    Set frm = VBA.UserForms.Add(pFormName)" & vbCrLf
    s = s & "    frm.Prepare lngId" & vbCrLf
    s = s & "    frm.Show" & vbCrLf
    s = s & "    If frm.Succeeded Then LoadList lngId" & vbCrLf
    s = s & "    Unload frm" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "EH:" & vbCrLf
    s = s & "    lblMessage.Caption = Err.Description" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdClose_Click()" & vbCrLf
    s = s & "    Me.Hide" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & vbCrLf
    s = s & "    If CloseMode = 0 Then" & vbCrLf
    s = s & "        Cancel = True" & vbCrLf
    s = s & "        Me.Hide" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    CodeFrmSales = s
End Function

Private Function CodeFrmSaleStart() As String
    Dim s As String

    s = s & "Option Explicit" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' KAMERIX - Nouvelle vente comptant (cree par ZmodFormBuilder : ne pas modifier a la main," & vbCrLf
    s = s & "' modifier le generateur puis relancer BuildKamerixForms). Ecriture : module Ventes." & vbCrLf
    s = s & "' La marge prevue (prix - cout de revient - commission) s'affiche pendant la saisie," & vbCrLf
    s = s & "' en rouge si elle est negative : la vente reste possible apres confirmation." & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private mblnSucceeded As Boolean" & vbCrLf
    s = s & "Private mlngSavedId As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Property Get Succeeded() As Boolean" & vbCrLf
    s = s & "    Succeeded = mblnSucceeded" & vbCrLf
    s = s & "End Property" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Property Get SavedId() As Long" & vbCrLf
    s = s & "    SavedId = mlngSavedId" & vbCrLf
    s = s & "End Property" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_Initialize()" & vbCrLf
    s = s & "    SetupCombo cboCar, 260" & vbCrLf
    s = s & "    SetupCombo cboBuyer, 200" & vbCrLf
    s = s & "    SetupCombo cboIntermediary, 200" & vbCrLf
    s = s & "    SetupCombo cboPlace, 200" & vbCrLf
    s = s & "    FillListControl cboCar, SaleableCarList(), Array(""VoitureID"", ""Libelle"")" & vbCrLf
    s = s & "    FillListControl cboBuyer, PartyList("""", PARTY_ROLE_BUYER), Array(""TiersID"", ""Nom"")" & vbCrLf
    s = s & "    FillListControl cboIntermediary, IntermediaryList(), Array(""IntermediaireID"", ""Nom"")" & vbCrLf
    s = s & "    AddNoneItem cboIntermediary" & vbCrLf
    s = s & "    cboIntermediary.ListIndex = 0" & vbCrLf
    s = s & "    FillListControl cboPlace, PlaceList(PLACE_TYPE_CITY), Array(""LieuID"", ""Nom"")" & vbCrLf
    s = s & "    AddNoneItem cboPlace" & vbCrLf
    s = s & "    cboPlace.ListIndex = 0" & vbCrLf
    s = s & "    txtCommission.Text = ""0""" & vbCrLf
    s = s & "    txtDate.Text = Format$(Date, ""dd/mm/yyyy"")" & vbCrLf
    s = s & "    lblMessage.Caption = """"" & vbCrLf
    s = s & "    If cboCar.ListCount = 0 Then lblMessage.Caption = ""Aucune voiture en stock avec un achat enregistr"" & ChrW(233) & "".""" & vbCrLf
    s = s & "    UpdateMargin" & vbCrLf
    s = s & "    cmdOK.Enabled = CanWrite(MOD_SALES)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub SetupCombo(ByVal pCombo As Object, ByVal pWidth As Single)" & vbCrLf
    s = s & "    pCombo.Style = 2" & vbCrLf
    s = s & "    pCombo.ColumnCount = 2" & vbCrLf
    s = s & "    pCombo.ColumnWidths = ""0 pt;"" & pWidth & "" pt""" & vbCrLf
    s = s & "    pCombo.TextColumn = 2" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cboCar_Change()" & vbCrLf
    s = s & "    UpdateMargin" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub txtPrice_Change()" & vbCrLf
    s = s & "    UpdateMargin" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub txtCommission_Change()" & vbCrLf
    s = s & "    UpdateMargin" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Cout de revient et marge prevue de la voiture choisie" & vbCrLf
    s = s & "Private Sub UpdateMargin()" & vbCrLf
    s = s & "    Dim curPrice As Currency" & vbCrLf
    s = s & "    Dim curCommission As Currency" & vbCrLf
    s = s & "    Dim curMargin As Currency" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo EH" & vbCrLf
    s = s & "    lblCost.Caption = """"" & vbCrLf
    s = s & "    lblMargin.Caption = """"" & vbCrLf
    s = s & "    If SelectedIdOf(cboCar) = 0 Then Exit Sub" & vbCrLf
    s = s & "    lblCost.Caption = ""Co"" & ChrW(251) & ""t de revient : "" & Format$(CarCostPrice(SelectedIdOf(cboCar)), ""#,##0"")" & vbCrLf
    s = s & "    If Not TryParseAmount(txtPrice.Text, curPrice) Or Not TryParseAmount(txtCommission.Text, curCommission) Then Exit Sub" & vbCrLf
    s = s & "    If curPrice = 0 Then Exit Sub" & vbCrLf
    s = s & "    curMargin = EstimatedMargin(SelectedIdOf(cboCar), curPrice, curCommission)" & vbCrLf
    s = s & "    lblMargin.Caption = ""Marge pr"" & ChrW(233) & ""vue : "" & Format$(curMargin, ""#,##0"")" & vbCrLf
    s = s & "    If curMargin < 0 Then lblMargin.ForeColor = RGB(192, 0, 0) Else lblMargin.ForeColor = RGB(0, 128, 0)" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "EH:" & vbCrLf
    s = s & "    lblCost.Caption = Err.Description" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdOK_Click()" & vbCrLf
    s = s & "    Dim curPrice As Currency" & vbCrLf
    s = s & "    Dim curCommission As Currency" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo EH" & vbCrLf
    s = s & "    lblMessage.Caption = """"" & vbCrLf
    s = s & "    If SelectedIdOf(cboCar) = 0 Or SelectedIdOf(cboBuyer) = 0 Then" & vbCrLf
    s = s & "        lblMessage.Caption = ""Choisissez la voiture et l'acheteur.""" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If Not TryParseAmount(txtPrice.Text, curPrice) Or Not TryParseAmount(txtCommission.Text, curCommission) Then" & vbCrLf
    s = s & "        lblMessage.Caption = ""Le prix et la commission doivent "" & ChrW(234) & ""tre des montants.""" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If Not IsDate(txtDate.Text) Then" & vbCrLf
    s = s & "        lblMessage.Caption = ""La date de la vente est invalide (jj/mm/aaaa).""" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If curPrice > 0 Then" & vbCrLf
    s = s & "        If EstimatedMargin(SelectedIdOf(cboCar), curPrice, curCommission) < 0 Then" & vbCrLf
    s = s & "            If MsgBox(""Ce prix est inf"" & ChrW(233) & ""rieur au co"" & ChrW(251) & ""t de revient : la vente sera "" & ChrW(224) & _" & vbCrLf
    s = s & "                      "" perte."" & vbCrLf & ""Continuer ?"", vbExclamation + vbYesNo, ""Kamerix"") <> vbYes Then Exit Sub" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    mlngSavedId = StartCashSale(SelectedIdOf(cboCar), SelectedIdOf(cboBuyer), curPrice, curCommission, _" & vbCrLf
    s = s & "                                SelectedIdOf(cboIntermediary), SelectedIdOf(cboPlace), CDate(txtDate.Text))" & vbCrLf
    s = s & "    mblnSucceeded = True" & vbCrLf
    s = s & "    Me.Hide" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "EH:" & vbCrLf
    s = s & "    lblMessage.Caption = Err.Description" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdCancel_Click()" & vbCrLf
    s = s & "    Me.Hide" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & vbCrLf
    s = s & "    If CloseMode = 0 Then" & vbCrLf
    s = s & "        Cancel = True" & vbCrLf
    s = s & "        Me.Hide" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    CodeFrmSaleStart = s
End Function

Private Function CodeFrmSalePayment() As String
    Dim s As String

    s = s & "Option Explicit" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' KAMERIX - Encaissement d'un paiement sur une vente (cree par ZmodFormBuilder : ne pas" & vbCrLf
    s = s & "' modifier a la main, modifier le generateur puis relancer BuildKamerixForms)." & vbCrLf
    s = s & "' Appeler Prepare (identifiant de la vente) avant Show. Ecriture : module Paiements." & vbCrLf
    s = s & "' Apres l'encaissement : impression du recu proposee, et de la facture si la vente est soldee." & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private mlngSaleId As Long" & vbCrLf
    s = s & "Private mblnSucceeded As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Property Get Succeeded() As Boolean" & vbCrLf
    s = s & "    Succeeded = mblnSucceeded" & vbCrLf
    s = s & "End Property" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_Initialize()" & vbCrLf
    s = s & "    Dim vModes As Variant" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    cboMode.Style = 2" & vbCrLf
    s = s & "    vModes = PaymentModes()" & vbCrLf
    s = s & "    For i = LBound(vModes) To UBound(vModes)" & vbCrLf
    s = s & "        cboMode.AddItem vModes(i)" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "    cboMode.ListIndex = 0" & vbCrLf
    s = s & "    txtDate.Text = Format$(Date, ""dd/mm/yyyy"")" & vbCrLf
    s = s & "    lblMessage.Caption = """"" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub Prepare(ByVal pSaleId As Long)" & vbCrLf
    s = s & "    Dim objContract As IContract" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    mlngSaleId = pSaleId" & vbCrLf
    s = s & "    Set objContract = LoadSaleContract(pSaleId)" & vbCrLf
    s = s & "    lblSale.Caption = SaleLabel(pSaleId)" & vbCrLf
    s = s & "    lblAmounts.Caption = ""Prix : "" & Format$(objContract.Price, ""#,##0"") & ""    Encaiss"" & ChrW(233) & "" : "" & _" & vbCrLf
    s = s & "                         Format$(objContract.AmountPaid, ""#,##0"") & ""    Solde : "" & Format$(objContract.Balance, ""#,##0"")" & vbCrLf
    s = s & "    txtAmount.Text = Format$(objContract.Balance, ""#,##0"")" & vbCrLf
    s = s & "    cmdOK.Enabled = CanWrite(MOD_PAYMENTS)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdOK_Click()" & vbCrLf
    s = s & "    Dim curAmount As Currency" & vbCrLf
    s = s & "    Dim strReceipt As String" & vbCrLf
    s = s & "    Dim strMessage As String" & vbCrLf
    s = s & "    Dim objInvoice As Invoice" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo EH" & vbCrLf
    s = s & "    lblMessage.Caption = """"" & vbCrLf
    s = s & "    If Not TryParseAmount(txtAmount.Text, curAmount) Then" & vbCrLf
    s = s & "        lblMessage.Caption = ""Le montant est invalide.""" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If Not IsDate(txtDate.Text) Then" & vbCrLf
    s = s & "        lblMessage.Caption = ""La date du paiement est invalide (jj/mm/aaaa).""" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    strReceipt = RecordSalePayment(mlngSaleId, curAmount, cboMode.Text, CDate(txtDate.Text))" & vbCrLf
    s = s & "    mblnSucceeded = True" & vbCrLf
    s = s & "    Me.Hide" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    strMessage = ""Paiement enregistr"" & ChrW(233) & "" : re"" & ChrW(231) & ""u "" & strReceipt & "".""" & vbCrLf
    s = s & "    Set objInvoice = InvoiceOfSale(mlngSaleId)" & vbCrLf
    s = s & "    If Not objInvoice Is Nothing Then" & vbCrLf
    s = s & "        If objInvoice.Status <> INVOICE_STATUS_ISSUED Then Set objInvoice = Nothing" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If Not objInvoice Is Nothing Then" & vbCrLf
    s = s & "        strMessage = strMessage & vbCrLf & ""La vente est sold"" & ChrW(233) & ""e : facture "" & objInvoice.InvoiceNumber & _" & vbCrLf
    s = s & "                     "" "" & ChrW(233) & ""mise ("" & Format$(objInvoice.TotalAmount, ""#,##0"") & "" TTC).""" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    OfferDocument DOC_RECEIPT, PaymentIdOfReceipt(strReceipt), strMessage, ""Imprimer le re"" & ChrW(231) & ""u maintenant ?""" & vbCrLf
    s = s & "    If Not objInvoice Is Nothing Then" & vbCrLf
    s = s & "        OfferDocument DOC_INVOICE, objInvoice.Id, ""Facture "" & objInvoice.InvoiceNumber & "" ("" & _" & vbCrLf
    s = s & "                      Format$(objInvoice.TotalAmount, ""#,##0"") & "" TTC)."", ""Imprimer la facture maintenant ?""" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "EH:" & vbCrLf
    s = s & "    lblMessage.Caption = Err.Description" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdCancel_Click()" & vbCrLf
    s = s & "    Me.Hide" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & vbCrLf
    s = s & "    If CloseMode = 0 Then" & vbCrLf
    s = s & "        Cancel = True" & vbCrLf
    s = s & "        Me.Hide" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    CodeFrmSalePayment = s
End Function

Private Function CodeFrmSaleCv() As String
    Dim s As String

    s = s & "Option Explicit" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' KAMERIX - Suivi du certificat de vente : depot au commissariat, puis signature (cree par" & vbCrLf
    s = s & "' ZmodFormBuilder : ne pas modifier a la main, modifier le generateur puis relancer" & vbCrLf
    s = s & "' BuildKamerixForms). Appeler Prepare (identifiant de la vente) avant Show." & vbCrLf
    s = s & "' Ecriture : module Ventes. Le depot exige une vente soldee ; la signature rend la voiture vendue." & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private mlngSaleId As Long" & vbCrLf
    s = s & "Private mblnSucceeded As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Property Get Succeeded() As Boolean" & vbCrLf
    s = s & "    Succeeded = mblnSucceeded" & vbCrLf
    s = s & "End Property" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_Initialize()" & vbCrLf
    s = s & "    cboStation.Style = 2" & vbCrLf
    s = s & "    cboStation.ColumnCount = 2" & vbCrLf
    s = s & "    cboStation.ColumnWidths = ""0 pt;180 pt""" & vbCrLf
    s = s & "    cboStation.TextColumn = 2" & vbCrLf
    s = s & "    FillListControl cboStation, PlaceList(PLACE_TYPE_POLICE_STATION), Array(""LieuID"", ""Nom"")" & vbCrLf
    s = s & "    lblMessage.Caption = """"" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub Prepare(ByVal pSaleId As Long)" & vbCrLf
    s = s & "    Dim objCv As CvTracking" & vbCrLf
    s = s & "    Dim blnWrite As Boolean" & vbCrLf
    s = s & "    Dim blnSettled As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    mlngSaleId = pSaleId" & vbCrLf
    s = s & "    lblSale.Caption = SaleLabel(pSaleId)" & vbCrLf
    s = s & "    Set objCv = CvOfSale(pSaleId)" & vbCrLf
    s = s & "    blnWrite = CanWrite(MOD_SALES)" & vbCrLf
    s = s & "    blnSettled = LoadSaleContract(pSaleId).IsSettled" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    lblCvStatus.Caption = ""Statut du CV : "" & CvStatusLabel(objCv.Status)" & vbCrLf
    s = s & "    If objCv.PoliceStationId <> 0 Then SelectById cboStation, objCv.PoliceStationId" & vbCrLf
    s = s & "    If IsNull(objCv.FilingDate) Then txtFilingDate.Text = Format$(Date, ""dd/mm/yyyy"") Else txtFilingDate.Text = Format$(objCv.FilingDate, ""dd/mm/yyyy"")" & vbCrLf
    s = s & "    txtFees.Text = Format$(objCv.FeesPaid, ""#,##0"")" & vbCrLf
    s = s & "    txtCvNumber.Text = objCv.CvNumber" & vbCrLf
    s = s & "    If IsNull(objCv.SignatureDate) Then txtSignatureDate.Text = Format$(Date, ""dd/mm/yyyy"") Else txtSignatureDate.Text = Format$(objCv.SignatureDate, ""dd/mm/yyyy"")" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    cmdSubmit.Enabled = blnWrite And objCv.Status = CV_STATUS_TO_FILE And blnSettled" & vbCrLf
    s = s & "    cmdSign.Enabled = blnWrite And objCv.Status = CV_STATUS_IN_PROGRESS" & vbCrLf
    s = s & "    cboStation.Enabled = cmdSubmit.Enabled" & vbCrLf
    s = s & "    txtFilingDate.Enabled = cmdSubmit.Enabled" & vbCrLf
    s = s & "    txtFees.Enabled = cmdSubmit.Enabled" & vbCrLf
    s = s & "    txtCvNumber.Enabled = cmdSign.Enabled" & vbCrLf
    s = s & "    txtSignatureDate.Enabled = cmdSign.Enabled" & vbCrLf
    s = s & "    If objCv.Status = CV_STATUS_TO_FILE And Not blnSettled Then" & vbCrLf
    s = s & "        lblMessage.Caption = ""Le dossier pourra "" & ChrW(234) & ""tre d"" & ChrW(233) & ""pos"" & ChrW(233) & "" quand la vente sera sold"" & ChrW(233) & ""e.""" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdSubmit_Click()" & vbCrLf
    s = s & "    Dim curFees As Currency" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo EH" & vbCrLf
    s = s & "    lblMessage.Caption = """"" & vbCrLf
    s = s & "    If SelectedIdOf(cboStation) = 0 Then" & vbCrLf
    s = s & "        lblMessage.Caption = ""Choisissez le commissariat.""" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If Not IsDate(txtFilingDate.Text) Or Not TryParseAmount(txtFees.Text, curFees) Then" & vbCrLf
    s = s & "        lblMessage.Caption = ""La date de d"" & ChrW(233) & ""p"" & ChrW(244) & ""t ou les frais sont invalides.""" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    SubmitCv mlngSaleId, SelectedIdOf(cboStation), CDate(txtFilingDate.Text), curFees" & vbCrLf
    s = s & "    mblnSucceeded = True" & vbCrLf
    s = s & "    Me.Hide" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "EH:" & vbCrLf
    s = s & "    lblMessage.Caption = Err.Description" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdSign_Click()" & vbCrLf
    s = s & "    On Error GoTo EH" & vbCrLf
    s = s & "    lblMessage.Caption = """"" & vbCrLf
    s = s & "    If Not IsDate(txtSignatureDate.Text) Then" & vbCrLf
    s = s & "        lblMessage.Caption = ""La date de signature est invalide (jj/mm/aaaa).""" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If MsgBox(""Enregistrer la signature du CV ?"" & vbCrLf & ""La voiture passera "" & ChrW(224) & "" Vendue (d"" & ChrW(233) & ""finitif)."", _" & vbCrLf
    s = s & "              vbQuestion + vbYesNo, ""Kamerix"") <> vbYes Then Exit Sub" & vbCrLf
    s = s & "    SignCv mlngSaleId, txtCvNumber.Text, CDate(txtSignatureDate.Text)" & vbCrLf
    s = s & "    mblnSucceeded = True" & vbCrLf
    s = s & "    Me.Hide" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "EH:" & vbCrLf
    s = s & "    lblMessage.Caption = Err.Description" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdClose_Click()" & vbCrLf
    s = s & "    Me.Hide" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & vbCrLf
    s = s & "    If CloseMode = 0 Then" & vbCrLf
    s = s & "        Cancel = True" & vbCrLf
    s = s & "        Me.Hide" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    CodeFrmSaleCv = s
End Function

Private Function CodeFrmSaleCancel() As String
    Dim s As String

    s = s & "Option Explicit" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' KAMERIX - Annulation d'une vente (cree par ZmodFormBuilder : ne pas modifier a la main," & vbCrLf
    s = s & "' modifier le generateur puis relancer BuildKamerixForms). Appeler Prepare avant Show." & vbCrLf
    s = s & "' Droits : VENTE_ANNULER, et REMBOURSER s'il y a un remboursement. Le remboursement" & vbCrLf
    s = s & "' affiche tient compte du desistement (frais du CV a la charge de l'acheteur)." & vbCrLf
    s = s & "' Apres l'annulation : impression du recu de remboursement proposee." & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private mlngSaleId As Long" & vbCrLf
    s = s & "Private mblnSucceeded As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Property Get Succeeded() As Boolean" & vbCrLf
    s = s & "    Succeeded = mblnSucceeded" & vbCrLf
    s = s & "End Property" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_Initialize()" & vbCrLf
    s = s & "    Dim vModes As Variant" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    cboMode.Style = 2" & vbCrLf
    s = s & "    vModes = PaymentModes()" & vbCrLf
    s = s & "    For i = LBound(vModes) To UBound(vModes)" & vbCrLf
    s = s & "        cboMode.AddItem vModes(i)" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "    cboMode.ListIndex = 0" & vbCrLf
    s = s & "    lblMessage.Caption = """"" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub Prepare(ByVal pSaleId As Long)" & vbCrLf
    s = s & "    mlngSaleId = pSaleId" & vbCrLf
    s = s & "    lblSale.Caption = SaleLabel(pSaleId)" & vbCrLf
    s = s & "    UpdateRefund" & vbCrLf
    s = s & "    cmdOK.Enabled = HasRight(RIGHT_CANCEL_SALE)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub chkWithdrawal_Click()" & vbCrLf
    s = s & "    UpdateRefund" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Montant rembourse selon le desistement" & vbCrLf
    s = s & "Private Sub UpdateRefund()" & vbCrLf
    s = s & "    Dim curRefund As Currency" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo EH" & vbCrLf
    s = s & "    curRefund = RefundAmountIfCancelled(mlngSaleId, chkWithdrawal.Value)" & vbCrLf
    s = s & "    lblRefund.Caption = ""Montant rembours"" & ChrW(233) & "" : "" & Format$(curRefund, ""#,##0"")" & vbCrLf
    s = s & "    cboMode.Enabled = (curRefund > 0)" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "EH:" & vbCrLf
    s = s & "    lblRefund.Caption = Err.Description" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdOK_Click()" & vbCrLf
    s = s & "    Dim strRefund As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo EH" & vbCrLf
    s = s & "    lblMessage.Caption = """"" & vbCrLf
    s = s & "    If Len(Trim$(txtReason.Text)) = 0 Then" & vbCrLf
    s = s & "        lblMessage.Caption = ""Le motif de l'annulation est obligatoire.""" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If MsgBox(""Annuler cette vente ?"" & vbCrLf & ""La voiture sera remise en stock."", vbQuestion + vbYesNo, ""Kamerix"") <> vbYes Then Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    strRefund = CancelSale(mlngSaleId, txtReason.Text, chkWithdrawal.Value, cboMode.Text)" & vbCrLf
    s = s & "    mblnSucceeded = True" & vbCrLf
    s = s & "    Me.Hide" & vbCrLf
    s = s & "    If Len(strRefund) > 0 Then" & vbCrLf
    s = s & "        OfferDocument DOC_REFUND, PaymentIdOfReceipt(strRefund), _" & vbCrLf
    s = s & "                      ""Vente annul"" & ChrW(233) & ""e. Remboursement enregistr"" & ChrW(233) & "" : pi"" & ChrW(232) & ""ce "" & strRefund & ""."", _" & vbCrLf
    s = s & "                      ""Imprimer le re"" & ChrW(231) & ""u de remboursement maintenant ?""" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        MsgBox ""Vente annul"" & ChrW(233) & ""e (rien "" & ChrW(224) & "" rembourser)."", vbInformation, ""Kamerix""" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "EH:" & vbCrLf
    s = s & "    lblMessage.Caption = Err.Description" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdCancel_Click()" & vbCrLf
    s = s & "    Me.Hide" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & vbCrLf
    s = s & "    If CloseMode = 0 Then" & vbCrLf
    s = s & "        Cancel = True" & vbCrLf
    s = s & "        Me.Hide" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    CodeFrmSaleCancel = s
End Function

Private Function CodeFrmSaleDocuments() As String
    Dim s As String

    s = s & "Option Explicit" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' KAMERIX - Documents d'une vente : facture, recus, remboursements (cree par ZmodFormBuilder :" & vbCrLf
    s = s & "' ne pas modifier a la main, modifier le generateur puis relancer BuildKamerixForms)." & vbCrLf
    s = s & "' Appeler Prepare (identifiant de la vente) avant Show. Lecture : module Ventes pour la liste ;" & vbCrLf
    s = s & "' Factures ou Paiements pour editer. Chaque edition cree un PDF (Documents\AAAA\...) ;" & vbCrLf
    s = s & "' a partir de la deuxieme, le document porte la mention DUPLICATA." & vbCrLf
    s = s & "' Colonnes de lstDocs : voir les constantes DOC_COL_* de modKamerixDocuments." & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private mlngSaleId As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Property Get Succeeded() As Boolean" & vbCrLf
    s = s & "    Succeeded = False" & vbCrLf
    s = s & "End Property" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_Initialize()" & vbCrLf
    s = s & "    lblMessage.Caption = """"" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub Prepare(ByVal pSaleId As Long)" & vbCrLf
    s = s & "    mlngSaleId = pSaleId" & vbCrLf
    s = s & "    lblSale.Caption = SaleLabel(pSaleId)" & vbCrLf
    s = s & "    LoadList 0" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub LoadList(ByVal pSelectRow As Long)" & vbCrLf
    s = s & "    Dim vData As Variant" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo EH" & vbCrLf
    s = s & "    lstDocs.Clear" & vbCrLf
    s = s & "    lstDocs.ColumnCount = 8" & vbCrLf
    s = s & "    lstDocs.ColumnWidths = ""0 pt;80 pt;105 pt;60 pt;80 pt;50 pt;70 pt;0 pt""" & vbCrLf
    s = s & "    vData = DocumentRows(mlngSaleId)" & vbCrLf
    s = s & "    If IsEmpty(vData) Then" & vbCrLf
    s = s & "        lblCount.Caption = ""Aucun document""" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        lstDocs.Column = vData" & vbCrLf
    s = s & "        lblCount.Caption = (UBound(vData, 2) + 1) & "" document(s)""" & vbCrLf
    s = s & "        If pSelectRow >= 0 And pSelectRow < lstDocs.ListCount Then lstDocs.ListIndex = pSelectRow" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    UpdateButtons" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "EH:" & vbCrLf
    s = s & "    ShowError Err.Description" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub lstDocs_Change()" & vbCrLf
    s = s & "    UpdateButtons" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UpdateButtons()" & vbCrLf
    s = s & "    Dim blnAllowed As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If lstDocs.ListIndex >= 0 Then blnAllowed = CanReadDocument(CStr(lstDocs.List(lstDocs.ListIndex, DOC_COL_TYPE)))" & vbCrLf
    s = s & "    cmdPrint.Enabled = blnAllowed" & vbCrLf
    s = s & "    cmdPdf.Enabled = blnAllowed" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdPrint_Click()" & vbCrLf
    s = s & "    Produce kxDocPrint" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdPdf_Click()" & vbCrLf
    s = s & "    Produce kxDocPdf" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub lstDocs_DblClick(ByVal Cancel As MSForms.ReturnBoolean)" & vbCrLf
    s = s & "    If cmdPdf.Enabled Then Produce kxDocPdf" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Edite le document selectionne (PDF ouvert, ou impression), puis recharge la liste" & vbCrLf
    s = s & "Private Sub Produce(ByVal pOutput As KxDocOutput)" & vbCrLf
    s = s & "    Dim lngRow As Long" & vbCrLf
    s = s & "    Dim strType As String" & vbCrLf
    s = s & "    Dim lngId As Long" & vbCrLf
    s = s & "    Dim strPath As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo EH" & vbCrLf
    s = s & "    lblMessage.Caption = """"" & vbCrLf
    s = s & "    lngRow = lstDocs.ListIndex" & vbCrLf
    s = s & "    If lngRow < 0 Then Exit Sub" & vbCrLf
    s = s & "    strType = CStr(lstDocs.List(lngRow, DOC_COL_TYPE))" & vbCrLf
    s = s & "    lngId = CLng(lstDocs.List(lngRow, DOC_COL_ID))" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Me.MousePointer = 11" & vbCrLf
    s = s & "    strPath = ProduceDocument(strType, lngId, pOutput)" & vbCrLf
    s = s & "    Me.MousePointer = 0" & vbCrLf
    s = s & "    LoadList lngRow" & vbCrLf
    s = s & "    lblMessage.ForeColor = RGB(0, 110, 0)" & vbCrLf
    s = s & "    If pOutput = kxDocPrint Then" & vbCrLf
    s = s & "        lblMessage.Caption = ""Document envoy"" & ChrW(233) & "" "" & ChrW(224) & "" l'imprimante. PDF : "" & strPath" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        lblMessage.Caption = ""PDF enregistr"" & ChrW(233) & "" : "" & strPath" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "EH:" & vbCrLf
    s = s & "    Me.MousePointer = 0" & vbCrLf
    s = s & "    ShowError Err.Description" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub ShowError(ByVal pText As String)" & vbCrLf
    s = s & "    lblMessage.ForeColor = RGB(192, 0, 0)" & vbCrLf
    s = s & "    lblMessage.Caption = pText" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdClose_Click()" & vbCrLf
    s = s & "    Me.Hide" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & vbCrLf
    s = s & "    If CloseMode = 0 Then" & vbCrLf
    s = s & "        Cancel = True" & vbCrLf
    s = s & "        Me.Hide" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    CodeFrmSaleDocuments = s
End Function

Private Function CodeFrmSettings() As String
    Dim s As String

    s = s & "Option Explicit" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' KAMERIX - Parametres (cree par ZmodFormBuilder : ne pas modifier a la main, modifier le" & vbCrLf
    s = s & "' generateur puis relancer BuildKamerixForms). Appeler Prepare avant Show :" & vbCrLf
    s = s & "'   Prepare True  : assistant de premiere configuration (onglet Entreprise seul) ;" & vbCrLf
    s = s & "'   Prepare False : ecran PARAMETRES (onglets Entreprise, Vente, Securite ; Enregistrer" & vbCrLf
    s = s & "'                   valide les trois onglets en une seule fois)." & vbCrLf
    s = s & "' Ecriture : module Administration. Logo de l'entreprise : copie dans le dossier Ressources a" & vbCrLf
    s = s & "' l'enregistrement ; apercu par LoadPictureGDIPlus (module ModulPictureGDIPlus, fond blanc)." & vbCrLf
    s = s & "' Logo de Kamerix (en haut a droite) : Images\AppLogo\Logo.png a cote du classeur, s'il existe." & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Const PAGE_COMPANY As Long = 0" & vbCrLf
    s = s & "Private Const PAGE_SALE As Long = 1" & vbCrLf
    s = s & "Private Const PAGE_SECURITY As Long = 2" & vbCrLf
    s = s & "Private Const APP_LOGO_PATH As String = ""\Images\AppLogo\Logo.png""" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private mblnSetup As Boolean" & vbCrLf
    s = s & "Private mblnSucceeded As Boolean" & vbCrLf
    s = s & "Private mstrLogoStored As String" & vbCrLf
    s = s & "Private mstrLogoSource As String" & vbCrLf
    s = s & "Private mblnLogoRemoved As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Property Get Succeeded() As Boolean" & vbCrLf
    s = s & "    Succeeded = mblnSucceeded" & vbCrLf
    s = s & "End Property" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_Initialize()" & vbCrLf
    s = s & "    lblMessage.Caption = """"" & vbCrLf
    s = s & "    lblLockInfo.Caption = """"" & vbCrLf
    s = s & "    txtModes.MultiLine = True" & vbCrLf
    s = s & "    txtModes.EnterKeyBehavior = True" & vbCrLf
    s = s & "    txtModes.ScrollBars = 2" & vbCrLf
    s = s & "    imgLogo.PictureSizeMode = 3" & vbCrLf
    s = s & "    imgAppLogo.PictureSizeMode = 3" & vbCrLf
    s = s & "    mpgSettings.Value = PAGE_COMPANY" & vbCrLf
    s = s & "    ShowAppLogo" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Logo de Kamerix (facultatif : rien ne s'affiche si le fichier est absent ou illisible)" & vbCrLf
    s = s & "Private Sub ShowAppLogo()" & vbCrLf
    s = s & "    Dim strPath As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    strPath = ThisWorkbook.Path & APP_LOGO_PATH" & vbCrLf
    s = s & "    If Len(Dir$(strPath)) = 0 Then Exit Sub" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    Set imgAppLogo.Picture = LoadPictureGDIPlus(strPath, GDIP_WHITE_BACKGROUND)" & vbCrLf
    s = s & "    On Error GoTo 0" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub Prepare(ByVal pSetup As Boolean)" & vbCrLf
    s = s & "    Dim objProfile As CompanyProfile" & vbCrLf
    s = s & "    Dim objSettings As GeneralSettings" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo EH" & vbCrLf
    s = s & "    mblnSetup = pSetup" & vbCrLf
    s = s & "    Set objProfile = LoadCompanyProfile()" & vbCrLf
    s = s & "    txtName.Text = objProfile.Name" & vbCrLf
    s = s & "    txtAddress.Text = objProfile.Address" & vbCrLf
    s = s & "    txtPhone.Text = objProfile.Phone" & vbCrLf
    s = s & "    txtEmail.Text = objProfile.Email" & vbCrLf
    s = s & "    txtNiu.Text = objProfile.Niu" & vbCrLf
    s = s & "    txtRccm.Text = objProfile.Rccm" & vbCrLf
    s = s & "    txtFooter.Text = objProfile.Footer" & vbCrLf
    s = s & "    mstrLogoStored = objProfile.LogoPath" & vbCrLf
    s = s & "    ShowLogo ResolveLogoPath(mstrLogoStored)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If pSetup Then" & vbCrLf
    s = s & "        Me.Caption = ""Kamerix - Configuration de l'entreprise""" & vbCrLf
    s = s & "        lblTitle.Caption = ""Bienvenue dans Kamerix""" & vbCrLf
    s = s & "        lblIntro.Caption = ""Renseignez votre entreprise : ces informations figurent sur les factures et les re"" & ChrW(231) & _" & vbCrLf
    s = s & "                           ""us. Les champs marqu"" & ChrW(233) & ""s * sont obligatoires. Vous pourrez les modifier plus tard "" & _" & vbCrLf
    s = s & "                           ""("" & ChrW(233) & ""cran PARAMETRES).""" & vbCrLf
    s = s & "        mpgSettings.Pages(PAGE_SALE).Visible = False" & vbCrLf
    s = s & "        mpgSettings.Pages(PAGE_SECURITY).Visible = False" & vbCrLf
    s = s & "        cmdOK.Caption = ""Enregistrer""" & vbCrLf
    s = s & "        cmdCancel.Caption = ""Plus tard""" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        lblIntro.Caption = ""Les champs marqu"" & ChrW(233) & ""s * sont obligatoires. Enregistrer valide les trois onglets.""" & vbCrLf
    s = s & "        Set objSettings = LoadGeneralSettings()" & vbCrLf
    s = s & "        txtTaxRate.Text = CStr(objSettings.TaxRate)" & vbCrLf
    s = s & "        txtModes.Text = Replace(objSettings.PaymentModesText, "";"", vbCrLf)" & vbCrLf
    s = s & "        txtCurrency.Text = objSettings.CurrencyCode" & vbCrLf
    s = s & "        txtDecimals.Text = CStr(objSettings.CurrencyDecimals)" & vbCrLf
    s = s & "        txtMaxFailures.Text = CStr(objSettings.MaxFailedLogins)" & vbCrLf
    s = s & "        txtLockMinutes.Text = CStr(objSettings.LockoutMinutes)" & vbCrLf
    s = s & "        If CurrencyLocked() Then" & vbCrLf
    s = s & "            txtCurrency.Enabled = False" & vbCrLf
    s = s & "            txtDecimals.Enabled = False" & vbCrLf
    s = s & "            lblLockInfo.Caption = ""Devise et d"" & ChrW(233) & ""cimales verrouill"" & ChrW(233) & ""es : des ventes sont d"" & _" & vbCrLf
    s = s & "                                  ChrW(233) & ""j"" & ChrW(224) & "" enregistr"" & ChrW(233) & ""es.""" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    cmdOK.Enabled = CanWrite(MOD_ADMINISTRATION)" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "EH:" & vbCrLf
    s = s & "    lblMessage.Caption = Err.Description" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' ---------- Logo ----------" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdBrowse_Click()" & vbCrLf
    s = s & "    Dim vFile As Variant" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    vFile = Application.GetOpenFilename(""Images (*.png;*.jpg;*.jpeg;*.bmp;*.gif),*.png;*.jpg;*.jpeg;*.bmp;*.gif"", , _" & vbCrLf
    s = s & "                                        ""Choisir le logo de l'entreprise"")" & vbCrLf
    s = s & "    If VarType(vFile) = vbBoolean Then Exit Sub" & vbCrLf
    s = s & "    mstrLogoSource = CStr(vFile)" & vbCrLf
    s = s & "    mblnLogoRemoved = False" & vbCrLf
    s = s & "    ShowLogo mstrLogoSource" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdRemoveLogo_Click()" & vbCrLf
    s = s & "    mstrLogoSource = """"" & vbCrLf
    s = s & "    mblnLogoRemoved = True" & vbCrLf
    s = s & "    ShowLogo """"" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Apercu : LoadPictureGDIPlus (tous formats, transparence sur fond blanc)" & vbCrLf
    s = s & "Private Sub ShowLogo(ByVal pPath As String)" & vbCrLf
    s = s & "    Dim objPicture As IPicture" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set imgLogo.Picture = Nothing" & vbCrLf
    s = s & "    If Len(pPath) = 0 Then" & vbCrLf
    s = s & "        lblLogoInfo.Caption = ""(aucun logo)""" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    lblLogoInfo.Caption = Mid$(pPath, InStrRev(pPath, ""\"") + 1)" & vbCrLf
    s = s & "    If Len(Dir$(pPath)) = 0 Then" & vbCrLf
    s = s & "        lblLogoInfo.Caption = lblLogoInfo.Caption & "" (fichier introuvable)""" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set objPicture = LoadPictureGDIPlus(pPath, GDIP_WHITE_BACKGROUND)" & vbCrLf
    s = s & "    If objPicture Is Nothing Then" & vbCrLf
    s = s & "        lblLogoInfo.Caption = lblLogoInfo.Caption & "" (aper"" & ChrW(231) & ""u indisponible)""" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        Set imgLogo.Picture = objPicture" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' ---------- Enregistrement ----------" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdOK_Click()" & vbCrLf
    s = s & "    Dim objProfile As CompanyProfile" & vbCrLf
    s = s & "    Dim objSettings As GeneralSettings" & vbCrLf
    s = s & "    Dim dblRate As Double" & vbCrLf
    s = s & "    Dim lngDecimals As Long" & vbCrLf
    s = s & "    Dim lngFailures As Long" & vbCrLf
    s = s & "    Dim lngMinutes As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error GoTo EH" & vbCrLf
    s = s & "    lblMessage.Caption = """"" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set objProfile = New CompanyProfile" & vbCrLf
    s = s & "    objProfile.Name = txtName.Text" & vbCrLf
    s = s & "    objProfile.Address = txtAddress.Text" & vbCrLf
    s = s & "    objProfile.Phone = txtPhone.Text" & vbCrLf
    s = s & "    objProfile.Email = txtEmail.Text" & vbCrLf
    s = s & "    objProfile.Niu = txtNiu.Text" & vbCrLf
    s = s & "    objProfile.Rccm = txtRccm.Text" & vbCrLf
    s = s & "    objProfile.Footer = txtFooter.Text" & vbCrLf
    s = s & "    If mblnLogoRemoved Then objProfile.LogoPath = """" Else objProfile.LogoPath = mstrLogoStored" & vbCrLf
    s = s & "    If Not NoErrors(objProfile.ValidationErrors(), PAGE_COMPANY) Then Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If Not mblnSetup Then" & vbCrLf
    s = s & "        If Not TryParseRate(txtTaxRate.Text, dblRate) Then" & vbCrLf
    s = s & "            ShowError ""Le taux de TVA est invalide (exemple : 19,25)."", PAGE_SALE" & vbCrLf
    s = s & "            Exit Sub" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "        If Not TryParseWholeNumber(txtDecimals.Text, lngDecimals) Then" & vbCrLf
    s = s & "            ShowError ""Le nombre de d"" & ChrW(233) & ""cimales est invalide."", PAGE_SALE" & vbCrLf
    s = s & "            Exit Sub" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "        If Not TryParseWholeNumber(txtMaxFailures.Text, lngFailures) Then" & vbCrLf
    s = s & "            ShowError ""Le nombre d'"" & ChrW(233) & ""checs avant blocage est invalide."", PAGE_SECURITY" & vbCrLf
    s = s & "            Exit Sub" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "        If Not TryParseWholeNumber(txtLockMinutes.Text, lngMinutes) Then" & vbCrLf
    s = s & "            ShowError ""La dur"" & ChrW(233) & ""e du blocage est invalide."", PAGE_SECURITY" & vbCrLf
    s = s & "            Exit Sub" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "        Set objSettings = New GeneralSettings" & vbCrLf
    s = s & "        objSettings.TaxRate = dblRate" & vbCrLf
    s = s & "        objSettings.PaymentModesText = txtModes.Text" & vbCrLf
    s = s & "        objSettings.CurrencyCode = txtCurrency.Text" & vbCrLf
    s = s & "        objSettings.CurrencyDecimals = lngDecimals" & vbCrLf
    s = s & "        objSettings.MaxFailedLogins = lngFailures" & vbCrLf
    s = s & "        objSettings.LockoutMinutes = lngMinutes" & vbCrLf
    s = s & "        If Not NoErrors(objSettings.SaleErrors(), PAGE_SALE) Then Exit Sub" & vbCrLf
    s = s & "        If Not NoErrors(objSettings.SecurityErrors(), PAGE_SECURITY) Then Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If MsgBox(""Enregistrer ces param"" & ChrW(232) & ""tres ?"", vbQuestion + vbYesNo, ""Kamerix"") <> vbYes Then Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Me.MousePointer = 11" & vbCrLf
    s = s & "    If Len(mstrLogoSource) > 0 Then objProfile.LogoPath = ImportLogo(mstrLogoSource)" & vbCrLf
    s = s & "    If mblnSetup Then" & vbCrLf
    s = s & "        SaveCompanyProfile objProfile" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        SaveAllSettings objProfile, objSettings" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Me.MousePointer = 0" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    mblnSucceeded = True" & vbCrLf
    s = s & "    Me.Hide" & vbCrLf
    s = s & "    If mblnSetup Then" & vbCrLf
    s = s & "        MsgBox ""Entreprise configur"" & ChrW(233) & ""e : les factures et les re"" & ChrW(231) & ""us peuvent maintenant "" & _" & vbCrLf
    s = s & "               ChrW(234) & ""tre "" & ChrW(233) & ""dit"" & ChrW(233) & ""s."", vbInformation, ""Kamerix""" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        MsgBox ""Param"" & ChrW(232) & ""tres enregistr"" & ChrW(233) & ""s."", vbInformation, ""Kamerix""" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "EH:" & vbCrLf
    s = s & "    Me.MousePointer = 0" & vbCrLf
    s = s & "    ShowError Err.Description, -1" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Vrai s'il n'y a aucune erreur ; sinon les affiche (2 au plus) et montre l'onglet concerne" & vbCrLf
    s = s & "Private Function NoErrors(ByVal pErrors As Collection, ByVal pPage As Long) As Boolean" & vbCrLf
    s = s & "    Dim strText As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If pErrors.Count = 0 Then" & vbCrLf
    s = s & "        NoErrors = True" & vbCrLf
    s = s & "        Exit Function" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    strText = pErrors.Item(1)" & vbCrLf
    s = s & "    If pErrors.Count > 1 Then strText = strText & vbCrLf & pErrors.Item(2)" & vbCrLf
    s = s & "    ShowError strText, pPage" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' Message d'erreur ; pPage >= 0 : affiche cet onglet" & vbCrLf
    s = s & "Private Sub ShowError(ByVal pText As String, ByVal pPage As Long)" & vbCrLf
    s = s & "    lblMessage.Caption = pText" & vbCrLf
    s = s & "    If pPage >= 0 Then" & vbCrLf
    s = s & "        If mpgSettings.Pages(pPage).Visible Then mpgSettings.Value = pPage" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdCancel_Click()" & vbCrLf
    s = s & "    Me.Hide" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & vbCrLf
    s = s & "    If CloseMode = 0 Then" & vbCrLf
    s = s & "        Cancel = True" & vbCrLf
    s = s & "        Me.Hide" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    CodeFrmSettings = s
End Function
