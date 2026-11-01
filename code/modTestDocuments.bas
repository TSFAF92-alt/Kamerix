Attribute VB_Name = "modTestDocuments"
Option Explicit
Option Private Module

'==============================================================================
' KAMERIX - Tests du lot 5.5 : montants en lettres, données et édition des documents
'------------------------------------------------------------------------------
' A lancer (F5 sur TestDocuments) sur une base de developpement SANS utilisateur.
' Tout se passe dans une transaction annulee a la fin : la base reste intacte.
' Les PDF sont crees dans un dossier temporaire (%TEMP%\KamerixTestDocs), supprime a la fin.
' Prerequis : une imprimante installee dans Windows (au besoin "Microsoft Print to PDF").
'==============================================================================

Private Const PWD_ADMIN As String = "Secret#2026"
Private Const TEST_TITLE As String = "Kamerix - tests des documents"

Public Sub TestDocuments()
    Dim lngOk As Long
    Dim lngTotal As Long
    Dim strReport As String
    Dim objData As DataAccess
    Dim objSeller As Party
    Dim objBuyer As Party
    Dim objInvoice As Invoice
    Dim d As Object
    Dim vRows As Variant
    Dim lngModel As Long
    Dim lngStation As Long
    Dim lngCarA As Long
    Dim lngCarB As Long
    Dim lngSaleA As Long
    Dim lngSaleB As Long
    Dim lngInvoiceA As Long
    Dim lngInvoiceB As Long
    Dim lngRec1 As Long
    Dim lngRec2 As Long
    Dim lngRefund As Long
    Dim strRec1 As String
    Dim strRec2 As String
    Dim strRefund As String
    Dim strPath As String
    Dim strPath2 As String
    Dim strFolder As String
    Dim curNet As Currency
    Dim curTax As Currency
    Dim lngErr As Long
    Dim strErrText As String
    Dim blnTransaction As Boolean

    On Error GoTo EH

    ' ---------- Montants en lettres (sans base) ----------
    CheckEq "Lettres : 0", AmountInWordsFor(0, "FCFA", 0), "z" & ChrW(233) & "ro franc CFA", lngOk, lngTotal, strReport
    CheckEq "Lettres : 1", AmountInWordsFor(1, "FCFA", 0), "un franc CFA", lngOk, lngTotal, strReport
    CheckEq "Lettres : 21 et 71 (et un, et onze)", NumberInWords(21) & " / " & NumberInWords(71), _
        "vingt et un / soixante et onze", lngOk, lngTotal, strReport
    CheckEq "Lettres : 80 et 81", NumberInWords(80) & " / " & NumberInWords(81), "quatre-vingts / quatre-vingt-un", _
        lngOk, lngTotal, strReport
    CheckEq "Lettres : 91 et 99", NumberInWords(91) & " / " & NumberInWords(99), _
        "quatre-vingt-onze / quatre-vingt-dix-neuf", lngOk, lngTotal, strReport
    CheckEq "Lettres : 200 et 201", NumberInWords(200) & " / " & NumberInWords(201), "deux cents / deux cent un", _
        lngOk, lngTotal, strReport
    CheckEq "Lettres : 80 000 et 200 000 (sans s devant mille)", NumberInWords(80000) & " / " & NumberInWords(200000), _
        "quatre-vingt mille / deux cent mille", lngOk, lngTotal, strReport
    CheckEq "Lettres : 1 000 et 21 000", NumberInWords(1000) & " / " & NumberInWords(21000), "mille / vingt et un mille", _
        lngOk, lngTotal, strReport
    CheckEq "Lettres : 5 000 000 (de francs)", AmountInWordsFor(5000000, "FCFA", 0), "cinq millions de francs CFA", _
        lngOk, lngTotal, strReport
    CheckEq "Lettres : 2 440 000", AmountInWordsFor(2440000, "FCFA", 0), _
        "deux millions quatre cent quarante mille francs CFA", lngOk, lngTotal, strReport
    CheckEq "Lettres : 280 000 000 (s devant millions)", AmountInWordsFor(280000000, "FCFA", 0), _
        "deux cent quatre-vingts millions de francs CFA", lngOk, lngTotal, strReport
    CheckEq "Lettres : 1 000 000 000", AmountInWordsFor(1000000000, "FCFA", 0), "un milliard de francs CFA", _
        lngOk, lngTotal, strReport
    CheckEq "Lettres : majuscule initiale", CapitalizeFirst("deux millions"), "Deux millions", lngOk, lngTotal, strReport
    CheckTrue "Lettres : montant negatif refuse", RaisesError(-1), lngOk, lngTotal, strReport

    ' ---------- Base de test ----------
    If CLng(QueryValue("SELECT COUNT(*) FROM [Utilisateur]")) > 0 Then
        MsgBox "Ce test demande une base sans utilisateur." & vbCrLf & _
               "Supprimez la base de developpement et recreez-la, puis relancez le test.", vbExclamation, TEST_TITLE
        Exit Sub
    End If

    strFolder = Environ$("TEMP") & Application.PathSeparator & "KamerixTestDocs"
    DeleteFolderTree strFolder
    SetDocumentFolderForTests strFolder

    CloseSession
    SetIterationsForTests 1000
    BeginTransaction
    blnTransaction = True

    CreateFirstAdmin "admin", "Administrateur Test", PWD_ADMIN
    OpenSession "admin", PWD_ADMIN

    EnsureDocumentSettings
    CheckTrue "Parametres de l'en-tete crees (SOCIETE_NOM, PIED_PAGE)", Len(CompanySetting(PARAM_COMPANY_NAME)) > 0 And _
        Len(CompanySetting(PARAM_FOOTER)) > 0, lngOk, lngTotal, strReport
    ExecuteNonQuery "UPDATE [Parametre] SET [Valeur] = ? WHERE [Cle] = ?", "TEST AUTO SARL", PARAM_COMPANY_NAME
    EnsureDocumentSettings
    CheckEq "Un parametre existant n'est jamais remplace", CompanySetting(PARAM_COMPANY_NAME), "TEST AUTO SARL", _
        lngOk, lngTotal, strReport

    lngModel = AddModel(AddBrand("TEST Toyota"), "Corolla")
    lngStation = AddPlace("TEST Commissariat", PLACE_TYPE_POLICE_STATION)
    Set objData = New DataAccess
    Set objSeller = New Party: objSeller.Name = "TEST Vendeur": objSeller.IsSeller = True: objData.SaveParty objSeller
    Set objBuyer = New Party
    objBuyer.Name = "TEST Client"
    objBuyer.IsBuyer = True
    objBuyer.Phone = "699000000"
    objBuyer.IdType = "CNI"
    objBuyer.IdNumber = "123456789"
    objData.SaveParty objBuyer

    lngCarA = PurchasedCar(lngModel, "TESTDOC-A", objSeller.Id, 3500000)
    lngCarB = PurchasedCar(lngModel, "TESTDOC-B", objSeller.Id, 2000000)

    ' Vente A : deux paiements, facture au solde
    lngSaleA = StartCashSale(lngCarA, objBuyer.Id, 5000000)
    CheckTrue "Vente sans paiement : aucun document", IsEmpty(DocumentRows(lngSaleA)), lngOk, lngTotal, strReport
    strRec1 = RecordSalePayment(lngSaleA, 2000000, "Esp" & ChrW(232) & "ces")
    strRec2 = RecordSalePayment(lngSaleA, 3000000, "Mobile Money")
    lngRec1 = PaymentIdOfReceipt(strRec1)
    lngRec2 = PaymentIdOfReceipt(strRec2)
    Set objInvoice = InvoiceOfSale(lngSaleA)
    lngInvoiceA = objInvoice.Id
    CheckTrue "Identifiants des pieces retrouves par leur numero", lngRec1 > 0 And lngRec2 > 0 And lngRec1 <> lngRec2 And _
        PaymentIdOfReceipt("REC-INCONNU") = 0, lngOk, lngTotal, strReport

    ' Vente B : soldee, CV depose (frais 60 000), annulee avec desistement -> remboursement 2 440 000
    lngSaleB = StartCashSale(lngCarB, objBuyer.Id, 2500000)
    RecordSalePayment lngSaleB, 2500000, "Esp" & ChrW(232) & "ces"
    lngInvoiceB = InvoiceOfSale(lngSaleB).Id
    SubmitCv lngSaleB, lngStation, Date, 60000
    strRefund = CancelSale(lngSaleB, "D" & ChrW(233) & "sistement du client", True, "Esp" & ChrW(232) & "ces")
    lngRefund = PaymentIdOfReceipt(strRefund)

    ' ---------- Donnees de la facture ----------
    Set d = DocumentData(DOC_INVOICE, lngInvoiceA)
    SplitTaxIncluded 5000000, curNet, curTax
    CheckTrue "Facture : numero, client et piece d'identite", d("Number") = objInvoice.InvoiceNumber And _
        d("PartyName") = "TEST Client" And d("PartyIdDoc") = "CNI n" & ChrW(176) & " 123456789", lngOk, lngTotal, strReport
    CheckTrue "Facture : HT + TVA = TTC = 5 000 000", d("Total") = 5000000 And d("Net") = curNet And d("Tax") = curTax And _
        d("Net") + d("Tax") = d("Total"), lngOk, lngTotal, strReport
    CheckEq "Facture : montant en lettres", d("AmountWords"), "cinq millions de francs CFA", lngOk, lngTotal, strReport
    vRows = d("PaymentRows")
    CheckTrue "Facture : 2 reglements, total regle 5 000 000", Not IsEmpty(vRows) And d("PaidTotal") = 5000000, _
        lngOk, lngTotal, strReport
    CheckTrue "Facture : reglements dans l'ordre (recus 1 puis 2)", ColumnValues(vRows, 0) = strRec1 & "|" & strRec2, _
        lngOk, lngTotal, strReport
    CheckEq "Facture : voiture (libelle avec l'annee, et chassis)", d("CarLabel") & " | " & d("Chassis"), _
        "TEST Toyota Corolla 2015 | TESTDOC-A", lngOk, lngTotal, strReport

    ' ---------- Donnees des recus ----------
    Set d = DocumentData(DOC_RECEIPT, lngRec1)
    CheckTrue "Recu 1 : encaisse 2 000 000, reste 3 000 000, pas de facture", d("Amount") = 2000000 And _
        d("PaidToDate") = 2000000 And d("Balance") = 3000000 And d("InvoiceNumber") = "", lngOk, lngTotal, strReport
    CheckEq "Recu 1 : mode", d("Mode"), "Esp" & ChrW(232) & "ces", lngOk, lngTotal, strReport
    Set d = DocumentData(DOC_RECEIPT, lngRec2)
    CheckTrue "Recu 2 (solde) : encaisse 5 000 000, reste 0, facture citee", d("PaidToDate") = 5000000 And _
        d("Balance") = 0 And d("InvoiceNumber") = objInvoice.InvoiceNumber, lngOk, lngTotal, strReport
    CheckEq "Recu 2 : montant en lettres", d("AmountWords"), "trois millions de francs CFA", lngOk, lngTotal, strReport

    ' ---------- Donnees du remboursement ----------
    Set d = DocumentData(DOC_REFUND, lngRefund)
    CheckTrue "Remboursement : verse 2 500 000, retenue 60 000, rembourse 2 440 000", d("PaidBefore") = 2500000 And _
        d("Withheld") = 60000 And d("Amount") = 2440000, lngOk, lngTotal, strReport
    CheckEq "Remboursement : motif", d("Reason"), "D" & ChrW(233) & "sistement du client", lngOk, lngTotal, strReport
    CheckTrue "Type controle : un encaissement n'est pas un remboursement", RaisesDocError(DOC_REFUND, lngRec1), _
        lngOk, lngTotal, strReport
    CheckTrue "Document introuvable : erreur", RaisesDocError(DOC_INVOICE, 999999), lngOk, lngTotal, strReport

    ' ---------- Liste des documents d'une vente ----------
    vRows = DocumentRows(lngSaleA)
    CheckEq "Vente A : 3 documents, facture puis recus 1 et 2 (types)", ColumnValues(vRows, DOC_COL_TYPE), _
        DOC_INVOICE & "|" & DOC_RECEIPT & "|" & DOC_RECEIPT, lngOk, lngTotal, strReport
    CheckEq "Vente A : numeros", ColumnValues(vRows, DOC_COL_NUMBER), objInvoice.InvoiceNumber & "|" & strRec1 & "|" & strRec2, _
        lngOk, lngTotal, strReport
    vRows = DocumentRows(lngSaleB)
    CheckTrue "Vente B : facture annulee signalee, remboursement liste", RowOfType(vRows, DOC_INVOICE) >= 0 And _
        RowOfType(vRows, DOC_REFUND) >= 0, lngOk, lngTotal, strReport
    If RowOfType(vRows, DOC_INVOICE) >= 0 Then
        CheckEq "Vente B : remarque de la facture", vRows(DOC_COL_NOTE, RowOfType(vRows, DOC_INVOICE)), "Annul" & ChrW(233) & "e", _
            lngOk, lngTotal, strReport
    End If

    ' ---------- Blocage tant que l'entreprise n'est pas configuree (lot 5.6) ----------
    CheckTrue "Entreprise non configuree : edition refusee", RaisesDocError(DOC_INVOICE, lngInvoiceA, True, ERR_DOCUMENT), _
        lngOk, lngTotal, strReport
    SaveCompanyProfile TestCompany()

    ' ---------- Edition des PDF ----------
    strPath = ProduceDocument(DOC_INVOICE, lngInvoiceA, kxDocFileOnly)
    CheckTrue "Facture : PDF cree (" & FileNameOf(strPath) & ")", FileExistsTest(strPath) And _
        FileNameOf(strPath) = objInvoice.InvoiceNumber & ".pdf", lngOk, lngTotal, strReport
    CheckTrue "Facture : rangee dans Documents\AAAA\Factures", InStr(strPath, Application.PathSeparator & _
        CStr(Year(Date)) & Application.PathSeparator & "Factures" & Application.PathSeparator) > 0, lngOk, lngTotal, strReport
    CheckTrue "Journal : 1 edition, original (Duplicata = Non)", EditionCount(DOC_INVOICE, lngInvoiceA) = 1 And _
        CLng(QueryValue("SELECT COUNT(*) FROM [JournalDocument] WHERE [TypeDocument] = ? AND [DocumentID] = ? AND [Duplicata] = False", _
                        DOC_INVOICE, lngInvoiceA)) = 1, lngOk, lngTotal, strReport

    strPath2 = ProduceDocument(DOC_INVOICE, lngInvoiceA, kxDocFileOnly)
    CheckTrue "Reedition : duplicata dans un nouveau fichier (_DUP2), original conserve", FileExistsTest(strPath2) And _
        InStr(FileNameOf(strPath2), "_DUP2") > 0 And FileExistsTest(strPath), lngOk, lngTotal, strReport
    CheckTrue "Journal : 2 editions dont 1 duplicata", EditionCount(DOC_INVOICE, lngInvoiceA) = 2 And _
        CLng(QueryValue("SELECT COUNT(*) FROM [JournalDocument] WHERE [TypeDocument] = ? AND [DocumentID] = ? AND [Duplicata] = True", _
                        DOC_INVOICE, lngInvoiceA)) = 1, lngOk, lngTotal, strReport

    strPath = ProduceDocument(DOC_RECEIPT, lngRec1, kxDocFileOnly)
    CheckTrue "Recu : PDF cree dans Recus", FileExistsTest(strPath) And InStr(strPath, Application.PathSeparator & "Recus" & _
        Application.PathSeparator) > 0, lngOk, lngTotal, strReport
    strPath = ProduceDocument(DOC_REFUND, lngRefund, kxDocFileOnly)
    CheckTrue "Remboursement : PDF cree dans Remboursements", FileExistsTest(strPath) And _
        InStr(strPath, "Remboursements") > 0, lngOk, lngTotal, strReport
    strPath = ProduceDocument(DOC_INVOICE, lngInvoiceB, kxDocFileOnly)
    CheckTrue "Facture annulee : PDF cree (mention ANNULEE)", FileExistsTest(strPath), lngOk, lngTotal, strReport

    vRows = DocumentRows(lngSaleA)
    If Not IsEmpty(vRows) Then
        CheckEq "Liste : nombre d'editions de la facture", vRows(DOC_COL_EDITIONS, 0), "2", lngOk, lngTotal, strReport
    End If

    ' ---------- Droits ----------
    CloseSession
    CheckTrue "Sans session : edition refusee", RaisesDocError(DOC_INVOICE, lngInvoiceA, True, ERR_SESSION_NOT_LOGGED_IN), _
        lngOk, lngTotal, strReport

    ' ---------- Fin : on annule tout ----------
    RollbackTransaction
    blnTransaction = False
    SetIterationsForTests 0
    SetDocumentFolderForTests ""
    DeleteFolderTree strFolder

    CheckEq "La base est restee intacte (aucun utilisateur)", CLng(QueryValue("SELECT COUNT(*) FROM [Utilisateur]")), 0, _
        lngOk, lngTotal, strReport
    CheckTrue "Dossier temporaire des PDF supprime", Len(Dir$(strFolder, vbDirectory)) = 0, lngOk, lngTotal, strReport

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
    SetDocumentFolderForTests ""
    DeleteFolderTree strFolder
    On Error GoTo 0
    Debug.Print "Echec (erreur " & lngErr & ") : " & strErrText & vbCrLf & strReport
    MsgBox "Echec (erreur " & lngErr & ") :" & vbCrLf & vbCrLf & strErrText & vbCrLf & vbCrLf & _
           "Resultats deja obtenus :" & vbCrLf & strReport, vbCritical, TEST_TITLE
End Sub

'==============================================================================
' OUTILS DU TEST
'==============================================================================

' Vrai si AmountInWordsFor leve une erreur pour ce montant
Private Function RaisesError(ByVal pAmount As Currency) As Boolean
    Dim strText As String

    On Error Resume Next
    strText = AmountInWordsFor(pAmount, "FCFA", 0)
    RaisesError = (Err.Number = ERR_AMOUNT_WORDS)
    On Error GoTo 0
End Function

' Vrai si DocumentData (ou ProduceDocument si pProduce) leve l'erreur attendue
' (par defaut ERR_DOCUMENT)
Private Function RaisesDocError(ByVal pDocType As String, ByVal pDocId As Long, Optional ByVal pProduce As Boolean = False, _
                                Optional ByVal pExpected As Long = ERR_DOCUMENT) As Boolean
    Dim d As Object
    Dim strPath As String

    On Error Resume Next
    If pProduce Then
        strPath = ProduceDocument(pDocType, pDocId, kxDocFileOnly)
    Else
        Set d = DocumentData(pDocType, pDocId)
    End If
    RaisesDocError = (Err.Number = pExpected)
    On Error GoTo 0
End Function

' Entreprise de test (configuration exigee avant toute edition)
Private Function TestCompany() As CompanyProfile
    Dim p As CompanyProfile

    Set p = New CompanyProfile
    p.Name = "TEST AUTO SARL"
    p.Address = "Akwa, Douala"
    p.Phone = "699000000"
    p.Footer = "Merci de votre confiance."
    Set TestCompany = p
End Function

' Valeurs d'une colonne d'un tableau (colonne, ligne) separees par | ; "" si Empty
Private Function ColumnValues(ByVal vData As Variant, ByVal pCol As Long) As String
    Dim i As Long

    If IsEmpty(vData) Then Exit Function
    For i = 0 To UBound(vData, 2)
        If i > 0 Then ColumnValues = ColumnValues & "|"
        ColumnValues = ColumnValues & CStr(vData(pCol, i))
    Next i
End Function

' Ligne du premier document de ce type ; -1 si absent
Private Function RowOfType(ByVal vData As Variant, ByVal pDocType As String) As Long
    Dim i As Long

    RowOfType = -1
    If IsEmpty(vData) Then Exit Function
    For i = 0 To UBound(vData, 2)
        If vData(DOC_COL_TYPE, i) = pDocType Then
            RowOfType = i
            Exit Function
        End If
    Next i
End Function

Private Function PurchasedCar(ByVal pModelId As Long, ByVal pChassis As String, ByVal pSellerId As Long, _
                              ByVal pPrice As Currency) As Long
    Dim objCar As Car

    Set objCar = New Car
    objCar.ModelId = pModelId
    objCar.ModelYear = 2015
    objCar.Color = "Grise"
    objCar.Chassis = pChassis
    objCar.Mileage = 85000
    RecordPurchase objCar, pSellerId, pPrice
    PurchasedCar = objCar.Id
End Function

Private Function FileNameOf(ByVal pPath As String) As String
    FileNameOf = Mid$(pPath, InStrRev(pPath, Application.PathSeparator) + 1)
End Function

Private Function FileExistsTest(ByVal pPath As String) As Boolean
    FileExistsTest = (Len(pPath) > 0) And (Len(Dir$(pPath)) > 0)
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

    ' Inventaire d'abord (Dir$ ne supporte pas les suppressions ni les appels imbriques)
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
