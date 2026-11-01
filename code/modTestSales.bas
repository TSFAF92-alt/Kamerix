Attribute VB_Name = "modTestSales"
Option Explicit
Option Private Module

'==============================================================================
' KAMERIX - Tests du lot 5.3 : vente comptant (paiements, facture, CV, annulation)
'------------------------------------------------------------------------------
' A lancer (F5 sur TestSales) sur une base de developpement SANS utilisateur,
' creee avec le schema du lot 5.1 (taux de taxe et modes de paiement).
' Tout se passe dans une transaction annulee a la fin : la base reste intacte.
'==============================================================================

Private Const PWD_ADMIN As String = "Secret#2026"
Private Const PWD_SELLER As String = "Vendeur#2026x"

Public Sub TestSales()
    Dim lngOk As Long
    Dim lngTotal As Long
    Dim strReport As String
    Dim objData As DataAccess
    Dim objSeller As Party
    Dim objBuyer As Party
    Dim objSellerOnly As Party
    Dim objInter As Intermediary
    Dim objContract As IContract
    Dim objInvoice As Invoice
    Dim objCv As CvTracking
    Dim lngSalesman As Long
    Dim lngModel As Long
    Dim lngCity As Long
    Dim lngStation As Long
    Dim lngCarA As Long
    Dim lngCarB As Long
    Dim lngCarC As Long
    Dim lngCarD As Long
    Dim lngSaleA As Long
    Dim lngSaleB As Long
    Dim lngSaleC As Long
    Dim strReceipt1 As String
    Dim strReceipt2 As String
    Dim strRefund As String
    Dim lngMovesBefore As Long
    Dim lngPaymentsBefore As Long
    Dim lngErr As Long
    Dim strErrText As String
    Dim blnTransaction As Boolean

    On Error GoTo EH

    If CLng(QueryValue("SELECT COUNT(*) FROM [Utilisateur]")) > 0 Then
        MsgBox "Ce test demande une base sans utilisateur." & vbCrLf & _
               "Supprimez la base de developpement et recreez-la, puis relancez le test.", _
               vbExclamation, "Kamerix - tests de la vente"
        Exit Sub
    End If

    lngMovesBefore = CLng(QueryValue("SELECT COUNT(*) FROM [Mouvement]"))
    lngPaymentsBefore = CLng(QueryValue("SELECT COUNT(*) FROM [Paiement]"))

    CloseSession
    SetIterationsForTests 1000
    BeginTransaction
    blnTransaction = True

    ' ---------- Preparation ----------
    CreateFirstAdmin "admin", "Administrateur Test", PWD_ADMIN
    lngSalesman = CreateUser("vendeur1", "Vendeur Test", GetRoleIdByName("Vendeur"), "Provisoire#1")
    ChangePassword lngSalesman, "Provisoire#1", PWD_SELLER
    OpenSession "admin", PWD_ADMIN

    lngModel = AddModel(AddBrand("TEST Toyota"), "Corolla")
    lngCity = AddPlace("TEST Douala", PLACE_TYPE_CITY)
    lngStation = AddPlace("TEST Commissariat central", PLACE_TYPE_POLICE_STATION, lngCity)

    Set objData = New DataAccess
    Set objSeller = NewParty("TEST Vendeur", False, True): objData.SaveParty objSeller
    Set objBuyer = NewParty("TEST Client", True, False): objData.SaveParty objBuyer
    Set objSellerOnly = NewParty("TEST Fournisseur", False, True): objData.SaveParty objSellerOnly
    Set objInter = New Intermediary
    objInter.Name = "TEST Demarcheur"
    objInter.CityId = lngCity
    objData.SaveIntermediary objInter

    lngCarA = PurchasedCar(objData, lngModel, "TESTVTE-A", objSeller.Id, 3500000, 150000, objInter.Id)   ' cout 3 650 000
    lngCarB = PurchasedCar(objData, lngModel, "TESTVTE-B", objSeller.Id, 2000000)
    lngCarC = PurchasedCar(objData, lngModel, "TESTVTE-C", objSeller.Id, 1500000)
    lngCarD = NewCarId(objData, lngModel, "TESTVTE-D")                                                     ' sans achat

    ' ---------- Demarrage de la vente ----------
    CheckEq "Voiture sans achat : vente refusee", SalesError("Start", lngCarD, objBuyer.Id, 3000000), ERR_SALE, _
        lngOk, lngTotal, strReport
    CheckEq "Tiers qui n'est pas acheteur refuse", SalesError("Start", lngCarA, objSellerOnly.Id, 5000000), ERR_SALE, _
        lngOk, lngTotal, strReport
    CheckEq "Marge prevue : 5 000 000 - 3 650 000 - 100 000 = 1 250 000", EstimatedMargin(lngCarA, 5000000, 100000), 1250000, _
        lngOk, lngTotal, strReport
    CheckEq "Marge prevue negative sous le cout (vente acceptee avec avertissement)", EstimatedMargin(lngCarA, 3000000), -650000, _
        lngOk, lngTotal, strReport

    OpenSession "vendeur1", PWD_SELLER
    lngSaleA = StartCashSale(lngCarA, objBuyer.Id, 5000000, 100000, objInter.Id, lngCity)
    CheckTrue "Le vendeur demarre la vente", lngSaleA > 0, lngOk, lngTotal, strReport
    CheckEq "Voiture en vente en cours", objData.LoadCar(lngCarA).Status, CAR_STATUS_SALE_IN_PROGRESS, lngOk, lngTotal, strReport
    CheckEq "Dossier de CV cree a deposer", objData.LoadCvTrackingOfMovement(lngSaleA).Status, CV_STATUS_TO_FILE, _
        lngOk, lngTotal, strReport
    CheckEq "Voiture deja en vente : deuxieme vente refusee", SalesError("Start", lngCarA, objBuyer.Id, 5000000), ERR_SALE, _
        lngOk, lngTotal, strReport
    CheckEq "Vendeur : paiement refuse (Paiements en lecture seule)", SalesError("Pay", lngSaleA, 1000000, "Esp" & ChrW(232) & "ces"), _
        ERR_SESSION_DENIED, lngOk, lngTotal, strReport
    CheckEq "Vendeur : annulation refusee (droit VENTE_ANNULER)", SalesError("Cancel", lngSaleA, False, "Esp" & ChrW(232) & "ces"), _
        ERR_SESSION_DENIED, lngOk, lngTotal, strReport

    OpenSession "admin", PWD_ADMIN
    Set objContract = LoadSaleContract(lngSaleA)
    CheckTrue "Contrat (IContract) : vente comptant", TypeOf objContract Is CashSale, lngOk, lngTotal, strReport
    CheckTrue "Contrat : prix 5 000 000, rien encaisse, solde 5 000 000", objContract.Price = 5000000 And _
        objContract.AmountPaid = 0 And objContract.Balance = 5000000 And Not objContract.IsSettled, lngOk, lngTotal, strReport

    ' ---------- Paiements ----------
    CheckEq "Paiement superieur au solde refuse", SalesError("Pay", lngSaleA, 6000000, "Esp" & ChrW(232) & "ces"), ERR_SALE, _
        lngOk, lngTotal, strReport
    CheckEq "Mode de paiement inconnu refuse", SalesError("Pay", lngSaleA, 1000000, "Bitcoin"), ERR_PAYMENT_INVALID, _
        lngOk, lngTotal, strReport
    CheckEq "Depot du CV avant le solde refuse", SalesError("SubmitCv", lngSaleA, lngStation), ERR_CV, lngOk, lngTotal, strReport

    strReceipt1 = RecordSalePayment(lngSaleA, 2000000, "Esp" & ChrW(232) & "ces")
    CheckTrue "Premier paiement : recu REC-annee-NNNNN", strReceipt1 Like "REC-" & Year(Date) & "-#####", lngOk, lngTotal, strReport
    CheckTrue "Avant le solde : pas de facture", InvoiceOfSale(lngSaleA) Is Nothing, lngOk, lngTotal, strReport
    CheckEq "Solde apres le premier paiement", LoadSaleContract(lngSaleA).Balance, 3000000, lngOk, lngTotal, strReport

    strReceipt2 = RecordSalePayment(lngSaleA, 3000000, "Mobile Money")
    CheckTrue "Second paiement : nouveau recu", strReceipt2 Like "REC-*" And strReceipt2 <> strReceipt1, lngOk, lngTotal, strReport
    CheckEq "Vente soldee au dernier paiement", objData.LoadMovement(lngSaleA).Status, MOVEMENT_STATUS_SETTLED, _
        lngOk, lngTotal, strReport

    Set objInvoice = InvoiceOfSale(lngSaleA)
    CheckTrue "Facture emise au solde (FAC-annee-NNNNN)", objInvoice.InvoiceNumber Like "FAC-" & Year(Date) & "-#####" And _
        objInvoice.Status = INVOICE_STATUS_ISSUED, lngOk, lngTotal, strReport
    CheckTrue "Facture : 5 000 000 TTC = 4 192 872 HT + 807 128 de taxe a 19,25 %", objInvoice.TotalAmount = 5000000 And _
        objInvoice.NetAmount = 4192872 And objInvoice.TaxAmount = 807128 And objInvoice.TaxRate = 19.25, lngOk, lngTotal, strReport
    CheckTrue "Facture : client et designation (modele, chassis)", objInvoice.PartyId = objBuyer.Id And _
        InStr(objInvoice.Designation, "TEST Toyota Corolla") > 0 And InStr(objInvoice.Designation, "TESTVTE-A") > 0, _
        lngOk, lngTotal, strReport
    CheckEq "Les 2 paiements sont rattaches a la facture", _
        CLng(QueryValue("SELECT COUNT(*) FROM [Paiement] WHERE [MouvementID] = ? AND [FactureID] = ?", lngSaleA, objInvoice.Id)), 2, _
        lngOk, lngTotal, strReport
    CheckEq "Vente soldee : paiement supplementaire refuse", SalesError("Pay", lngSaleA, 1000, "Esp" & ChrW(232) & "ces"), ERR_SALE, _
        lngOk, lngTotal, strReport

    ' ---------- Certificat de vente ----------
    CheckEq "Depot du CV dans un lieu qui n'est pas un commissariat refuse", SalesError("SubmitCv", lngSaleA, lngCity), ERR_CV, _
        lngOk, lngTotal, strReport
    SubmitCv lngSaleA, lngStation, Date, 75000
    Set objCv = CvOfSale(lngSaleA)
    CheckTrue "CV depose : en cours, commissariat et frais", objCv.Status = CV_STATUS_IN_PROGRESS And _
        objCv.PoliceStationId = lngStation And objCv.FeesPaid = 75000, lngOk, lngTotal, strReport
    CheckEq "Marge : 5 000 000 - 3 650 000 - 100 000 - 75 000 = 1 175 000", SaleMargin(lngSaleA), 1175000, lngOk, lngTotal, strReport

    SignCv lngSaleA, "CV-2026-123", Date
    CheckTrue "CV signe avec son numero", CvOfSale(lngSaleA).Status = CV_STATUS_SIGNED And _
        CvOfSale(lngSaleA).CvNumber = "CV-2026-123", lngOk, lngTotal, strReport
    CheckEq "Voiture vendue (definitif)", objData.LoadCar(lngCarA).Status, CAR_STATUS_SOLD, lngOk, lngTotal, strReport
    CheckEq "CV signe : annulation refusee", SalesError("Cancel", lngSaleA, False, "Esp" & ChrW(232) & "ces"), ERR_SALE, _
        lngOk, lngTotal, strReport

    ' ---------- Annulation sans desistement : remboursement total ----------
    lngSaleB = StartCashSale(lngCarB, objBuyer.Id, 2500000)
    RecordSalePayment lngSaleB, 1000000, "Esp" & ChrW(232) & "ces"
    CheckEq "Remboursement prevu : totalite versee", RefundAmountIfCancelled(lngSaleB, False), 1000000, lngOk, lngTotal, strReport
    CheckEq "Annulation sans mode de remboursement refusee", SalesError("Cancel", lngSaleB, False, ""), ERR_PAYMENT_INVALID, _
        lngOk, lngTotal, strReport
    CheckTrue "Refus : rien n'a change (vente en cours, voiture en vente)", _
        objData.LoadMovement(lngSaleB).Status = MOVEMENT_STATUS_IN_PROGRESS And _
        objData.LoadCar(lngCarB).Status = CAR_STATUS_SALE_IN_PROGRESS, lngOk, lngTotal, strReport

    strRefund = CancelSale(lngSaleB, "Financement du client refuse", False, "Esp" & ChrW(232) & "ces")
    CheckTrue "Annulation : piece de remboursement RMB-", strRefund Like "RMB-" & Year(Date) & "-#####", lngOk, lngTotal, strReport
    CheckTrue "Annulation : vente annulee avec motif, voiture remise en stock", _
        objData.LoadMovement(lngSaleB).Status = MOVEMENT_STATUS_CANCELLED And _
        objData.LoadMovement(lngSaleB).CancelReason = "Financement du client refuse" And _
        objData.LoadCar(lngCarB).Status = CAR_STATUS_IN_STOCK, lngOk, lngTotal, strReport
    CheckEq "Remboursement de 1 000 000 enregistre", objData.PaymentsTotal(lngSaleB, PAYMENT_DIRECTION_REFUND), 1000000, _
        lngOk, lngTotal, strReport
    CheckEq "Vente deja annulee : nouvelle annulation refusee", SalesError("Cancel", lngSaleB, False, "Esp" & ChrW(232) & "ces"), _
        ERR_SALE, lngOk, lngTotal, strReport

    ' ---------- Desistement de l'acheteur : frais du CV a sa charge ----------
    lngSaleC = StartCashSale(lngCarC, objBuyer.Id, 2000000)
    RecordSalePayment lngSaleC, 2000000, "Virement bancaire"
    SubmitCv lngSaleC, lngStation, Date, 60000
    CheckEq "Desistement : remboursement prevu = verse - frais du CV", RefundAmountIfCancelled(lngSaleC, True), 1940000, _
        lngOk, lngTotal, strReport
    CheckEq "Autre cause : remboursement prevu = totalite", RefundAmountIfCancelled(lngSaleC, False), 2000000, _
        lngOk, lngTotal, strReport
    CancelSale lngSaleC, "Desistement de l'acheteur", True, "Virement bancaire"
    CheckEq "Desistement : 1 940 000 rembourses", objData.PaymentsTotal(lngSaleC, PAYMENT_DIRECTION_REFUND), 1940000, _
        lngOk, lngTotal, strReport
    CheckEq "Annulation d'une vente soldee : facture annulee", objData.LoadInvoiceOfMovement(lngSaleC).Status, _
        INVOICE_STATUS_CANCELLED, lngOk, lngTotal, strReport
    CheckTrue "La voiture remise en stock peut etre revendue", StartCashSale(lngCarC, objBuyer.Id, 2100000) > 0, _
        lngOk, lngTotal, strReport

    ' ---------- Fin : on annule tout ----------
    CloseSession
    RollbackTransaction
    blnTransaction = False
    SetIterationsForTests 0

    CheckEq "La base est restee intacte (mouvements)", CLng(QueryValue("SELECT COUNT(*) FROM [Mouvement]")), lngMovesBefore, _
        lngOk, lngTotal, strReport
    CheckEq "La base est restee intacte (paiements)", CLng(QueryValue("SELECT COUNT(*) FROM [Paiement]")), lngPaymentsBefore, _
        lngOk, lngTotal, strReport
    CheckEq "La base est restee intacte (aucun utilisateur)", CLng(QueryValue("SELECT COUNT(*) FROM [Utilisateur]")), 0, _
        lngOk, lngTotal, strReport

    Debug.Print lngOk & " tests reussis sur " & lngTotal & "." & vbCrLf & strReport   ' rapport complet (Ctrl+G)
    MsgBox lngOk & " tests reussis sur " & lngTotal & "." & vbCrLf & vbCrLf & strReport, _
        IIf(lngOk = lngTotal, vbInformation, vbExclamation), "Kamerix - tests de la vente"
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
           "Resultats deja obtenus :" & vbCrLf & strReport, vbCritical, "Kamerix - tests de la vente"
End Sub

' Numero d'erreur d'une operation de vente (0 si acceptee)
Private Function SalesError(ByVal pOperation As String, ByVal pId As Long, Optional ByVal pArg1 As Variant, _
                            Optional ByVal pArg2 As Variant, Optional ByVal pArg3 As Variant) As Long
    On Error Resume Next
    Select Case pOperation
        Case "Start":    StartCashSale pId, CLng(pArg1), CCur(pArg2)
        Case "Pay":      RecordSalePayment pId, CCur(pArg1), CStr(pArg2)
        Case "SubmitCv": SubmitCv pId, CLng(pArg1), Date, 50000
        Case "Cancel":   CancelSale pId, "Test d'annulation", CBool(pArg1), CStr(pArg2)
    End Select
    SalesError = Err.Number
    Err.Clear
End Function

' Voiture creee et achetee ; retourne son identifiant
Private Function PurchasedCar(ByVal pData As DataAccess, ByVal pModelId As Long, ByVal pChassis As String, _
                              ByVal pSellerId As Long, ByVal pPrice As Currency, _
                              Optional ByVal pCommission As Currency = 0, Optional ByVal pIntermediaryId As Long = 0) As Long
    Dim objCar As Car

    Set objCar = New Car
    objCar.ModelId = pModelId
    objCar.Chassis = pChassis
    objCar.ModelYear = 2015
    RecordPurchase objCar, pSellerId, pPrice, pCommission, pIntermediaryId
    PurchasedCar = objCar.Id
End Function

' Voiture enregistree sans achat
Private Function NewCarId(ByVal pData As DataAccess, ByVal pModelId As Long, ByVal pChassis As String) As Long
    Dim objCar As Car

    Set objCar = New Car
    objCar.ModelId = pModelId
    objCar.Chassis = pChassis
    pData.SaveCar objCar
    NewCarId = objCar.Id
End Function

Private Function NewParty(ByVal pName As String, ByVal pIsBuyer As Boolean, ByVal pIsSeller As Boolean) As Party
    Set NewParty = New Party
    NewParty.Name = pName
    NewParty.IsBuyer = pIsBuyer
    NewParty.IsSeller = pIsSeller
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
