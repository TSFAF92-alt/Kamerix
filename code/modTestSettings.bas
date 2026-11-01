Attribute VB_Name = "modTestSettings"
Option Explicit
Option Private Module

'==============================================================================
' KAMERIX - Tests du lot 5.6 : paramétrage de l'entreprise et réglages généraux
'------------------------------------------------------------------------------
' A lancer (F5 sur TestSettings, ou TestSettings dans la fenetre Execution) sur une
' base de developpement SANS utilisateur. Tout se passe dans une transaction annulee
' a la fin : la base reste intacte. Le logo de test est copie dans
' %TEMP%\KamerixTestSettings, supprime a la fin.
'==============================================================================

Private Const PWD_ADMIN As String = "Secret#2026"
Private Const TEST_TITLE As String = "Kamerix - tests des parametres"

Public Sub TestSettings()
    Dim lngOk As Long
    Dim lngTotal As Long
    Dim strReport As String
    Dim objProfile As CompanyProfile
    Dim objSettings As GeneralSettings
    Dim objData As DataAccess
    Dim objSeller As Party
    Dim objBuyer As Party
    Dim dblRate As Double
    Dim lngNumber As Long
    Dim strFolder As String
    Dim strSource As String
    Dim strStored As String
    Dim lngModel As Long
    Dim lngCar As Long
    Dim lngErr As Long
    Dim strErrText As String
    Dim blnTransaction As Boolean

    On Error GoTo EH

    ' ---------- Saisie (sans base) ----------
    CheckTrue "Taux : 19,25 et 19.25 acceptes", TryParseRate("19,25", dblRate) And dblRate = 19.25 And _
        TryParseRate(" 19.25 ", dblRate) And dblRate = 19.25, lngOk, lngTotal, strReport
    CheckTrue "Taux : texte, negatif et double separateur refuses", Not TryParseRate("abc", dblRate) And _
        Not TryParseRate("-1", dblRate) And Not TryParseRate("19,2,5", dblRate) And Not TryParseRate("", dblRate), _
        lngOk, lngTotal, strReport
    CheckTrue "Entier : 15 accepte ; 5,5 et vide refuses", TryParseWholeNumber("15", lngNumber) And lngNumber = 15 And _
        Not TryParseWholeNumber("5,5", lngNumber) And Not TryParseWholeNumber("", lngNumber), lngOk, lngTotal, strReport

    ' ---------- CompanyProfile (sans base) ----------
    Set objProfile = New CompanyProfile
    CheckEq "Entreprise vide : 3 erreurs (nom, adresse, telephone)", objProfile.ValidationErrors().Count, 3, _
        lngOk, lngTotal, strReport
    Set objProfile = ValidProfile()
    CheckEq "Entreprise valide : aucune erreur", objProfile.ValidationErrors().Count, 0, lngOk, lngTotal, strReport
    CheckEq "NIU et RCCM : espaces retires, majuscules", objProfile.Niu & " | " & objProfile.Rccm, _
        "M012345678901X | RC/DLA/2026/B/1234", lngOk, lngTotal, strReport
    objProfile.Phone = "12ab"
    CheckEq "Telephone invalide refuse", objProfile.ValidationErrors().Count, 1, lngOk, lngTotal, strReport
    Set objProfile = ValidProfile()
    objProfile.Email = "contact@kamerix"
    CheckEq "E-mail sans point apres @ refuse", objProfile.ValidationErrors().Count, 1, lngOk, lngTotal, strReport
    objProfile.Email = ""
    CheckEq "E-mail facultatif", objProfile.ValidationErrors().Count, 0, lngOk, lngTotal, strReport
    objProfile.Footer = String$(151, "x")
    CheckEq "Pied de page de plus de 150 caracteres refuse", objProfile.ValidationErrors().Count, 1, lngOk, lngTotal, strReport

    ' ---------- GeneralSettings (sans base) ----------
    Set objSettings = ValidSettings()
    CheckEq "Reglages valides : aucune erreur", objSettings.ValidationErrors().Count, 0, lngOk, lngTotal, strReport
    objSettings.PaymentModesText = " Esp" & ChrW(232) & "ces " & vbCrLf & vbCrLf & "Mobile Money" & vbCrLf & "Ch" & ChrW(232) & "que" & vbCrLf
    CheckEq "Modes saisis un par ligne : normalises avec ;", objSettings.PaymentModesText, _
        "Esp" & ChrW(232) & "ces;Mobile Money;Ch" & ChrW(232) & "que", lngOk, lngTotal, strReport
    objSettings.TaxRate = 100
    CheckEq "TVA de 100 % refusee (onglet Vente)", objSettings.SaleErrors().Count, 1, lngOk, lngTotal, strReport
    Set objSettings = ValidSettings()
    objSettings.PaymentModesText = ""
    CheckEq "Aucun mode de paiement : refuse", objSettings.SaleErrors().Count, 1, lngOk, lngTotal, strReport
    objSettings.PaymentModesText = "Paiement par carte bancaire;Esp" & ChrW(232) & "ces;ESP" & ChrW(200) & "CES"
    CheckEq "Mode trop long et mode en double : 2 erreurs", objSettings.SaleErrors().Count, 2, lngOk, lngTotal, strReport
    Set objSettings = ValidSettings()
    objSettings.CurrencyDecimals = 5
    objSettings.CurrencyCode = ""
    CheckEq "Devise vide et 5 decimales : 2 erreurs", objSettings.SaleErrors().Count, 2, lngOk, lngTotal, strReport
    Set objSettings = ValidSettings()
    objSettings.MaxFailedLogins = 0
    objSettings.LockoutMinutes = 2000
    CheckTrue "Securite hors limites : 2 erreurs (onglet Securite), aucune sur Vente", _
        objSettings.SecurityErrors().Count = 2 And objSettings.SaleErrors().Count = 0, lngOk, lngTotal, strReport

    ' ---------- Base de test ----------
    If CLng(QueryValue("SELECT COUNT(*) FROM [Utilisateur]")) > 0 Then
        MsgBox "Ce test demande une base sans utilisateur." & vbCrLf & _
               "Supprimez la base de developpement et recreez-la, puis relancez le test.", vbExclamation, TEST_TITLE
        Exit Sub
    End If

    strFolder = Environ$("TEMP") & Application.PathSeparator & "KamerixTestSettings"
    DeleteFolderTree strFolder
    MkDir strFolder
    SetCompanyFolderForTests strFolder

    CloseSession
    SetIterationsForTests 1000
    BeginTransaction
    blnTransaction = True

    CheckTrue "Sans session : enregistrement refuse", RaisesError(ERR_SESSION_NOT_LOGGED_IN, ValidProfile()), _
        lngOk, lngTotal, strReport

    CreateFirstAdmin "admin", "Administrateur Test", PWD_ADMIN
    OpenSession "admin", PWD_ADMIN
    EnsureDocumentSettings

    ' ---------- Entreprise ----------
    CheckTrue "Base neuve : entreprise non configuree", Not IsCompanyConfigured(), lngOk, lngTotal, strReport
    Set objProfile = LoadCompanyProfile()
    CheckTrue "Non configuree : valeurs d'exemple NON proposees (nom vide, pied de page par defaut)", _
        objProfile.Name = "" And objProfile.Footer = "Merci de votre confiance.", lngOk, lngTotal, strReport
    Set objProfile = New CompanyProfile
    CheckTrue "Entreprise invalide : refusee, configuration non faite", RaisesError(ERR_COMPANY_INVALID, objProfile) And _
        Not IsCompanyConfigured(), lngOk, lngTotal, strReport

    SaveCompanyProfile ValidProfile()
    CheckTrue "Entreprise enregistree : configuration faite", IsCompanyConfigured(), lngOk, lngTotal, strReport
    CheckEq "Parametres ecrits (nom, telephone, NIU)", CompanySetting(PARAM_COMPANY_NAME) & " | " & _
        CompanySetting(PARAM_COMPANY_PHONE) & " | " & CompanySetting(PARAM_COMPANY_NIU), _
        "TS AUTO PRO | +237 699 00 00 00 | M012345678901X", lngOk, lngTotal, strReport
    Set objProfile = LoadCompanyProfile()
    CheckEq "Relecture de l'entreprise", objProfile.Name & " | " & objProfile.Address & " | " & objProfile.Email, _
        "TS AUTO PRO | Akwa, Douala | contact@tsauto.cm", lngOk, lngTotal, strReport

    ' ---------- Logo ----------
    strSource = strFolder & Application.PathSeparator & "source_logo.png"
    WriteTestFile strSource
    strStored = ImportLogo(strSource)
    CheckTrue "Logo copie dans Ressources, chemin relatif stocke (" & strStored & ")", _
        Left$(strStored, Len(RESOURCE_FOLDER_NAME) + 1) = RESOURCE_FOLDER_NAME & Application.PathSeparator And _
        Right$(strStored, 4) = ".png", lngOk, lngTotal, strReport
    CheckTrue "Chemin complet du logo retrouve, fichier present", Len(Dir$(ResolveLogoPath(strStored))) > 0 And _
        Left$(ResolveLogoPath(strStored), Len(strFolder)) = strFolder, lngOk, lngTotal, strReport
    CheckEq "Chemin absolu conserve tel quel", ResolveLogoPath("D:\Logos\a.png"), "D:\Logos\a.png", lngOk, lngTotal, strReport
    WriteTestFile strFolder & Application.PathSeparator & "notes.txt"
    CheckTrue "Logo : format .txt refuse", RaisesLogoError(strFolder & Application.PathSeparator & "notes.txt"), _
        lngOk, lngTotal, strReport
    CheckTrue "Logo : fichier introuvable refuse", RaisesLogoError(strFolder & Application.PathSeparator & "absent.png"), _
        lngOk, lngTotal, strReport

    ' ---------- Reglages generaux ----------
    Set objSettings = LoadGeneralSettings()
    CheckTrue "Reglages lus : TVA 19,25, FCFA, 0 decimale, 5 echecs, 15 minutes", objSettings.TaxRate = 19.25 And _
        objSettings.CurrencyCode = "FCFA" And objSettings.CurrencyDecimals = 0 And objSettings.MaxFailedLogins = 5 And _
        objSettings.LockoutMinutes = 15, lngOk, lngTotal, strReport

    Set objProfile = ValidProfile()
    objProfile.LogoPath = strStored
    Set objSettings = ValidSettings()
    objSettings.TaxRate = 18.5
    objSettings.PaymentModesText = "Esp" & ChrW(232) & "ces;Mobile Money;Ch" & ChrW(232) & "que"
    objSettings.MaxFailedLogins = 3
    objSettings.LockoutMinutes = 30
    SaveAllSettings objProfile, objSettings
    CheckTrue "Nouveaux reglages appliques (TVA 18,5 ; 3 modes ; Cheque accepte)", TaxRate() = 18.5 And _
        UBound(PaymentModes()) = 2 And IsValidPaymentMode("Ch" & ChrW(232) & "que"), lngOk, lngTotal, strReport
    CheckTrue "TVA stockee avec un point decimal (18.5)", CompanySetting(PARAM_TAX_RATE) = "18.5", lngOk, lngTotal, strReport
    CheckTrue "Securite appliquee (3 echecs, 30 minutes)", MaxFailedAttempts() = 3 And LockoutMinutes() = 30, _
        lngOk, lngTotal, strReport
    CheckEq "Logo enregistre avec les reglages", CompanySetting(PARAM_COMPANY_LOGO), strStored, lngOk, lngTotal, strReport

    ' ---------- Verrou de la devise ----------
    CheckTrue "Aucune vente : devise modifiable", Not CurrencyLocked(), lngOk, lngTotal, strReport
    lngModel = AddModel(AddBrand("TEST Toyota"), "Corolla")
    Set objData = New DataAccess
    Set objSeller = New Party: objSeller.Name = "TEST Vendeur": objSeller.IsSeller = True: objData.SaveParty objSeller
    Set objBuyer = New Party: objBuyer.Name = "TEST Client": objBuyer.IsBuyer = True: objData.SaveParty objBuyer
    lngCar = PurchasedCar(lngModel, "TESTPAR-A", objSeller.Id, 3000000)
    StartCashSale lngCar, objBuyer.Id, 4000000
    CheckTrue "Une vente existe : devise verrouillee", CurrencyLocked(), lngOk, lngTotal, strReport

    Set objSettings = ValidSettings()
    objSettings.CurrencyCode = "EUR"
    CheckTrue "Devise modifiee apres une vente : refusee, FCFA conserve", RaisesSettingsError(objProfile, objSettings) And _
        CurrencyCode() = "FCFA", lngOk, lngTotal, strReport
    Set objSettings = ValidSettings()
    objSettings.CurrencyDecimals = 2
    CheckTrue "Decimales modifiees apres une vente : refusees", RaisesSettingsError(objProfile, objSettings) And _
        CurrencyDecimals() = 0, lngOk, lngTotal, strReport
    Set objSettings = ValidSettings()
    objSettings.TaxRate = 19.25
    SaveAllSettings objProfile, objSettings
    CheckTrue "Apres une vente : TVA toujours modifiable (19,25)", TaxRate() = 19.25, lngOk, lngTotal, strReport

    ' ---------- Fin : on annule tout ----------
    CloseSession
    RollbackTransaction
    blnTransaction = False
    SetIterationsForTests 0
    SetCompanyFolderForTests ""
    DeleteFolderTree strFolder

    CheckEq "La base est restee intacte (aucun utilisateur)", CLng(QueryValue("SELECT COUNT(*) FROM [Utilisateur]")), 0, _
        lngOk, lngTotal, strReport
    CheckTrue "Dossier temporaire supprime", Len(Dir$(strFolder, vbDirectory)) = 0, lngOk, lngTotal, strReport

    Debug.Print lngOk & " tests reussis sur " & lngTotal & "." & vbCrLf & strReport   ' rapport complet (Ctrl+G)
    MsgBox lngOk & " tests reussis sur " & lngTotal & "." & vbCrLf & vbCrLf & strReport, _
        IIf(lngOk = lngTotal, vbInformation, vbExclamation), TEST_TITLE
    Exit Sub

EH:
    lngErr = Err.Number
    strErrText = Err.Description
    On Error Resume Next
    CloseSession
    If blnTransaction Then RollbackTransaction
    SetIterationsForTests 0
    SetCompanyFolderForTests ""
    DeleteFolderTree strFolder
    On Error GoTo 0
    Debug.Print "Echec (erreur " & lngErr & ") : " & strErrText & vbCrLf & strReport
    MsgBox "Echec (erreur " & lngErr & ") :" & vbCrLf & vbCrLf & strErrText & vbCrLf & vbCrLf & _
           "Resultats deja obtenus :" & vbCrLf & strReport, vbCritical, TEST_TITLE
End Sub

'==============================================================================
' OUTILS DU TEST
'==============================================================================

Private Function ValidProfile() As CompanyProfile
    Dim p As CompanyProfile

    Set p = New CompanyProfile
    p.Name = "TS AUTO PRO"
    p.Address = "Akwa, Douala"
    p.Phone = "+237 699 00 00 00"
    p.Email = "contact@tsauto.cm"
    p.Niu = " m012345678901x "
    p.Rccm = "rc/dla/2026/b/1234"
    p.Footer = "Merci de votre confiance."
    Set ValidProfile = p
End Function

Private Function ValidSettings() As GeneralSettings
    Dim g As GeneralSettings

    Set g = New GeneralSettings
    g.TaxRate = 19.25
    g.PaymentModesText = "Esp" & ChrW(232) & "ces;Mobile Money;Virement bancaire"
    g.CurrencyCode = "FCFA"
    g.CurrencyDecimals = 0
    g.MaxFailedLogins = 5
    g.LockoutMinutes = 15
    Set ValidSettings = g
End Function

' Vrai si SaveCompanyProfile leve l'erreur attendue
Private Function RaisesError(ByVal pExpected As Long, ByVal pProfile As CompanyProfile) As Boolean
    On Error Resume Next
    SaveCompanyProfile pProfile
    RaisesError = (Err.Number = pExpected)
    On Error GoTo 0
End Function

' Vrai si SaveAllSettings leve ERR_SETTINGS_INVALID
Private Function RaisesSettingsError(ByVal pProfile As CompanyProfile, ByVal pSettings As GeneralSettings) As Boolean
    On Error Resume Next
    SaveAllSettings pProfile, pSettings
    RaisesSettingsError = (Err.Number = ERR_SETTINGS_INVALID)
    On Error GoTo 0
End Function

' Vrai si ImportLogo leve ERR_LOGO
Private Function RaisesLogoError(ByVal pPath As String) As Boolean
    Dim strStored As String

    On Error Resume Next
    strStored = ImportLogo(pPath)
    RaisesLogoError = (Err.Number = ERR_LOGO)
    On Error GoTo 0
End Function

' Petit fichier de test (son contenu importe peu : ImportLogo ne fait que le copier)
Private Sub WriteTestFile(ByVal pPath As String)
    Dim intFile As Integer

    intFile = FreeFile
    Open pPath For Output As #intFile
    Print #intFile, "Kamerix - fichier de test"
    Close #intFile
End Sub

Private Function PurchasedCar(ByVal pModelId As Long, ByVal pChassis As String, ByVal pSellerId As Long, _
                              ByVal pPrice As Currency) As Long
    Dim objCar As Car

    Set objCar = New Car
    objCar.ModelId = pModelId
    objCar.Chassis = pChassis
    RecordPurchase objCar, pSellerId, pPrice
    PurchasedCar = objCar.Id
End Function

' Supprime un dossier, ses fichiers et ses sous-dossiers (sans erreur s'il n'existe pas)
Private Sub DeleteFolderTree(ByVal pFolder As String)
    Dim colFolders As Collection
    Dim colFiles As Collection
    Dim strName As String
    Dim strPath As String
    Dim i As Long

    If Len(pFolder) = 0 Then Exit Sub
    If Len(Dir$(pFolder, vbDirectory)) = 0 Then Exit Sub

    Set colFolders = New Collection
    Set colFiles = New Collection
    strName = Dir$(pFolder & Application.PathSeparator & "*", vbDirectory)
    Do While Len(strName) > 0
        If strName <> "." And strName <> ".." Then
            strPath = pFolder & Application.PathSeparator & strName
            If (GetAttr(strPath) And vbDirectory) = vbDirectory Then colFolders.Add strPath Else colFiles.Add strPath
        End If
        strName = Dir$()
    Loop

    For i = 1 To colFiles.Count
        Kill CStr(colFiles(i))
    Next i
    For i = 1 To colFolders.Count
        DeleteFolderTree CStr(colFolders(i))
    Next i
    RmDir pFolder
End Sub

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
