Attribute VB_Name = "modKamerixCompany"
Option Explicit
Option Private Module

'==============================================================================
' KAMERIX - Paramétrage de l'entreprise et des réglages généraux (lot 5.6)
'------------------------------------------------------------------------------
' Rôle        : charger et enregistrer l'identité de l'entreprise (CompanyProfile)
'               et les réglages de vente et de sécurité (GeneralSettings), saisis
'               dans frmSettings (assistant de première configuration ou écran
'               PARAMETRES). Rien ne se modifie dans Access.
' Configuration : CONFIG_ENTREPRISE_FAITE = 1 après le premier enregistrement
'               valide de l'entreprise. Tant qu'elle n'est pas faite, l'édition
'               des documents est refusée (modKamerixDocuments.ProduceDocument).
' Logo        : ImportLogo copie l'image choisie dans le dossier Ressources, à
'               côté du classeur, et retourne le chemin RELATIF à stocker
'               (Ressources\logo_aaaammjj_hhnnss.png) : Kamerix peut être déplacé
'               sans perdre son logo. ResolveLogoPath retrouve le chemin complet.
' Garde-fous  : devise et décimales verrouillées dès qu'une vente existe (les
'               montants passés changeraient de sens) ; un nouveau taux de TVA ne
'               touche que les futures factures (chaque facture garde le sien).
' Droits      : écriture sur le module Administration pour enregistrer ; lecture
'               sur Administration pour charger.
' Dépendances : CompanyProfile, GeneralSettings, modKamerixSql, modKamerixSession,
'               modKamerixSettings, modKamerixSecurity, modKamerixDocuments.
'==============================================================================

Public Const ERR_COMPANY_INVALID As Long = vbObjectError + 1660     ' données de l'entreprise invalides
Public Const ERR_SETTINGS_INVALID As Long = vbObjectError + 1661    ' réglages de vente ou de sécurité invalides
Public Const ERR_LOGO As Long = vbObjectError + 1662                ' logo introuvable ou refusé

Public Const PARAM_COMPANY_CONFIGURED As String = "CONFIG_ENTREPRISE_FAITE"
Public Const RESOURCE_FOLDER_NAME As String = "Ressources"

Private Const LOGO_EXTENSIONS As String = "|png|jpg|jpeg|bmp|gif|"
Private Const MAX_LOGO_BYTES As Long = 5242880                       ' 5 Mo
Private Const DEFAULT_FOOTER As String = "Merci de votre confiance."

' Dossier racine imposé par les tests à la place du dossier du classeur ("" = normal)
Private mstrTestRoot As String

'==============================================================================
' ETAT DE LA CONFIGURATION
'==============================================================================

' Vrai quand l'entreprise a été configurée et validée au moins une fois
Public Function IsCompanyConfigured() As Boolean
    IsCompanyConfigured = (CompanySetting(PARAM_COMPANY_CONFIGURED) = "1")
End Function

' Vrai dès qu'une vente existe (même annulée) : devise et décimales ne se modifient plus
Public Function CurrencyLocked() As Boolean
    CurrencyLocked = CLng(QueryValue("SELECT COUNT(*) FROM [Mouvement] WHERE [TypeMouvement] = ?", MOVEMENT_TYPE_SALE)) > 0
End Function

'==============================================================================
' ENTREPRISE
'==============================================================================

' Identité de l'entreprise. Tant qu'elle n'est pas configurée, les valeurs d'exemple
' ne sont PAS proposées : champs vides (pied de page par défaut). Lecture : Administration.
Public Function LoadCompanyProfile() As CompanyProfile
    Dim p As CompanyProfile

    RequireRead MOD_ADMINISTRATION
    Set p = New CompanyProfile
    If IsCompanyConfigured() Then
        p.Name = CompanySetting(PARAM_COMPANY_NAME)
        p.Address = CompanySetting(PARAM_COMPANY_ADDRESS)
        p.Phone = CompanySetting(PARAM_COMPANY_PHONE)
        p.Email = CompanySetting(PARAM_COMPANY_EMAIL)
        p.Niu = CompanySetting(PARAM_COMPANY_NIU)
        p.Rccm = CompanySetting(PARAM_COMPANY_RCCM)
        p.LogoPath = CompanySetting(PARAM_COMPANY_LOGO)
        p.Footer = CompanySetting(PARAM_FOOTER)
    Else
        p.Footer = DEFAULT_FOOTER
    End If
    Set LoadCompanyProfile = p
End Function

' Enregistre l'entreprise et marque la configuration comme faite. Écriture : Administration.
Public Sub SaveCompanyProfile(ByVal pProfile As CompanyProfile)
    Dim blnOwn As Boolean
    Dim lngErr As Long
    Dim strErr As String
    Dim strSrc As String

    RequireWrite MOD_ADMINISTRATION
    pProfile.Validate

    blnOwn = BeginIfNeeded()
    On Error GoTo EH
    WriteCompanyProfile pProfile
    CommitIfOwn blnOwn
    Exit Sub

EH:
    lngErr = Err.Number: strErr = Err.Description: strSrc = Err.Source
    RollbackIfOwn blnOwn
    Err.Raise lngErr, strSrc, strErr
End Sub

'==============================================================================
' REGLAGES GENERAUX (VENTE ET SECURITE)
'==============================================================================

' Réglages actuels lus dans la base. Lecture : Administration.
Public Function LoadGeneralSettings() As GeneralSettings
    Dim g As GeneralSettings

    RequireRead MOD_ADMINISTRATION
    Set g = New GeneralSettings
    g.TaxRate = TaxRate()
    g.PaymentModesText = Join(PaymentModes(), ";")
    g.CurrencyCode = CurrencyCode()
    g.CurrencyDecimals = CurrencyDecimals()
    g.MaxFailedLogins = MaxFailedAttempts()
    g.LockoutMinutes = LockoutMinutes()
    Set LoadGeneralSettings = g
End Function

' Enregistre l'entreprise ET les réglages généraux en une seule transaction (écran
' PARAMETRES : le bouton Enregistrer valide les trois onglets). Écriture : Administration.
Public Sub SaveAllSettings(ByVal pProfile As CompanyProfile, ByVal pSettings As GeneralSettings)
    Const SRC As String = "modKamerixCompany.SaveAllSettings"
    Dim blnOwn As Boolean
    Dim lngErr As Long
    Dim strErr As String
    Dim strSrc As String

    RequireWrite MOD_ADMINISTRATION
    pProfile.Validate
    pSettings.Validate
    If CurrencyLocked() Then
        If StrComp(pSettings.CurrencyCode, CurrencyCode(), vbTextCompare) <> 0 Or _
           pSettings.CurrencyDecimals <> CurrencyDecimals() Then
            Err.Raise ERR_SETTINGS_INVALID, SRC, "Des ventes sont d" & ChrW(233) & "j" & ChrW(224) & " enregistr" & ChrW(233) & _
                "es : la devise et le nombre de d" & ChrW(233) & "cimales ne peuvent plus " & ChrW(234) & "tre modifi" & ChrW(233) & "s."
        End If
    End If

    blnOwn = BeginIfNeeded()
    On Error GoTo EH
    WriteCompanyProfile pProfile
    SetSetting PARAM_TAX_RATE, Trim$(Str$(pSettings.TaxRate))          ' Str$ : toujours un point décimal
    SetSetting PARAM_PAYMENT_MODES, pSettings.PaymentModesText
    SetSetting PARAM_CURRENCY, pSettings.CurrencyCode
    SetSetting PARAM_CURRENCY_DECIMALS, CStr(pSettings.CurrencyDecimals)
    SetSetting PARAM_MAX_FAILURES, CStr(pSettings.MaxFailedLogins)
    SetSetting PARAM_LOCKOUT_MINUTES, CStr(pSettings.LockoutMinutes)
    CommitIfOwn blnOwn
    Exit Sub

EH:
    lngErr = Err.Number: strErr = Err.Description: strSrc = Err.Source
    RollbackIfOwn blnOwn
    Err.Raise lngErr, strSrc, strErr
End Sub

'==============================================================================
' LOGO
'==============================================================================

' Copie l'image choisie dans le dossier Ressources et retourne le chemin relatif à
' stocker dans SOCIETE_LOGO. Formats : png, jpg, jpeg, bmp, gif ; 5 Mo au plus.
Public Function ImportLogo(ByVal pSourcePath As String) As String
    Const SRC As String = "modKamerixCompany.ImportLogo"
    Dim strExt As String
    Dim strName As String
    Dim strFolder As String

    pSourcePath = Trim$(pSourcePath)
    If Not FileExists(pSourcePath) Then
        Err.Raise ERR_LOGO, SRC, "Fichier du logo introuvable : " & pSourcePath
    End If
    strExt = LCase$(Mid$(pSourcePath, InStrRev(pSourcePath, ".") + 1))
    If InStrRev(pSourcePath, ".") = 0 Or InStr(LOGO_EXTENSIONS, "|" & strExt & "|") = 0 Then
        Err.Raise ERR_LOGO, SRC, "Format de logo refus" & ChrW(233) & " (." & strExt & ") : PNG, JPG, BMP ou GIF."
    End If
    If FileLen(pSourcePath) > MAX_LOGO_BYTES Then
        Err.Raise ERR_LOGO, SRC, "Logo trop volumineux (5 Mo au plus)."
    End If

    strFolder = AppRoot() & Application.PathSeparator & RESOURCE_FOLDER_NAME
    EnsureFolderTree strFolder
    strName = "logo_" & Format$(Now, "yyyymmdd_hhnnss") & "." & strExt
    FileCopy pSourcePath, strFolder & Application.PathSeparator & strName

    ImportLogo = RESOURCE_FOLDER_NAME & Application.PathSeparator & strName
End Function

' Chemin complet d'un logo stocké : relatif -> dossier de Kamerix + chemin ; absolu -> tel quel
Public Function ResolveLogoPath(ByVal pStoredPath As String) As String
    pStoredPath = Trim$(pStoredPath)
    If Len(pStoredPath) = 0 Then Exit Function
    If Mid$(pStoredPath, 2, 1) = ":" Or Left$(pStoredPath, 2) = "\\" Then
        ResolveLogoPath = pStoredPath
    Else
        ResolveLogoPath = AppRoot() & Application.PathSeparator & pStoredPath
    End If
End Function

'==============================================================================
' SAISIE (fonctions pures, utilisées par frmSettings et testées sans formulaire)
'==============================================================================

' Taux saisi "19,25" ou "19.25" -> 19.25 ; Faux si ce n'est pas un nombre positif ou nul
Public Function TryParseRate(ByVal pText As String, ByRef pRate As Double) As Boolean
    Dim strValue As String

    strValue = Replace(Replace(Trim$(pText), ",", "."), " ", "")
    If Len(strValue) = 0 Or strValue = "." Then Exit Function
    If strValue Like "*[!0-9.]*" Then Exit Function
    If Len(strValue) - Len(Replace(strValue, ".", "")) > 1 Then Exit Function
    pRate = Val(strValue)
    TryParseRate = True
End Function

' Nombre entier positif ou nul saisi ("15") ; Faux sinon
Public Function TryParseWholeNumber(ByVal pText As String, ByRef pNumber As Long) As Boolean
    Dim strValue As String

    strValue = Trim$(pText)
    If Len(strValue) = 0 Or Len(strValue) > 9 Then Exit Function
    If strValue Like "*[!0-9]*" Then Exit Function
    pNumber = CLng(strValue)
    TryParseWholeNumber = True
End Function

'==============================================================================
' OUTILS DE TEST
'==============================================================================

' Dossier racine utilisé à la place du dossier du classeur pendant les tests ("" = normal)
Public Sub SetCompanyFolderForTests(ByVal pFolder As String)
    mstrTestRoot = pFolder
End Sub

'==============================================================================
' PROCEDURES INTERNES
'==============================================================================

' Écrit les paramètres de l'entreprise et marque la configuration (dans la transaction de l'appelant)
Private Sub WriteCompanyProfile(ByVal pProfile As CompanyProfile)
    SetSetting PARAM_COMPANY_NAME, pProfile.Name
    SetSetting PARAM_COMPANY_ADDRESS, pProfile.Address
    SetSetting PARAM_COMPANY_PHONE, pProfile.Phone
    SetSetting PARAM_COMPANY_EMAIL, pProfile.Email
    SetSetting PARAM_COMPANY_NIU, pProfile.Niu
    SetSetting PARAM_COMPANY_RCCM, pProfile.Rccm
    SetSetting PARAM_COMPANY_LOGO, pProfile.LogoPath
    SetSetting PARAM_FOOTER, pProfile.Footer
    SetSetting PARAM_COMPANY_CONFIGURED, "1"
End Sub

' Met à jour un paramètre, ou le crée s'il n'existe pas (texte vide -> Null)
Private Sub SetSetting(ByVal pKey As String, ByVal pValue As String)
    If CLng(QueryValue("SELECT COUNT(*) FROM [Parametre] WHERE [Cle] = ?", pKey)) = 0 Then
        ExecuteNonQuery "INSERT INTO [Parametre] ([Cle], [Valeur]) VALUES (?, ?)", pKey, NullIfEmpty(pValue)
    Else
        ExecuteNonQuery "UPDATE [Parametre] SET [Valeur] = ? WHERE [Cle] = ?", NullIfEmpty(pValue), pKey
    End If
End Sub

' Dossier de Kamerix (celui du classeur), ou le dossier imposé par les tests
Private Function AppRoot() As String
    If Len(mstrTestRoot) > 0 Then
        AppRoot = mstrTestRoot
    Else
        If Len(ThisWorkbook.Path) = 0 Then
            Err.Raise ERR_LOGO, "modKamerixCompany", "Enregistrez d'abord le classeur Kamerix."
        End If
        AppRoot = ThisWorkbook.Path
    End If
End Function

' Crée le dossier et ses parents manquants
Private Sub EnsureFolderTree(ByVal pFolder As String)
    Dim lngPos As Long

    If FolderExists(pFolder) Then Exit Sub
    If Right$(pFolder, 1) = ":" Then Exit Sub
    lngPos = InStrRev(pFolder, Application.PathSeparator)
    If lngPos > 1 Then EnsureFolderTree Left$(pFolder, lngPos - 1)
    MkDir pFolder
End Sub

Private Function FolderExists(ByVal pFolder As String) As Boolean
    On Error Resume Next
    FolderExists = ((GetAttr(pFolder) And vbDirectory) = vbDirectory)
    On Error GoTo 0
End Function

Private Function FileExists(ByVal pPath As String) As Boolean
    If Len(pPath) = 0 Then Exit Function
    On Error Resume Next
    FileExists = ((GetAttr(pPath) And vbDirectory) = 0)
    If Err.Number <> 0 Then FileExists = False
    On Error GoTo 0
End Function

' Transaction : ouverte ici seulement si l'appelant n'en a pas déjà une
Private Function BeginIfNeeded() As Boolean
    If Not TransactionActive() Then
        BeginTransaction
        BeginIfNeeded = True
    End If
End Function

Private Sub CommitIfOwn(ByVal pOwn As Boolean)
    If pOwn Then CommitTransaction
End Sub

Private Sub RollbackIfOwn(ByVal pOwn As Boolean)
    If pOwn Then RollbackTransaction
End Sub
