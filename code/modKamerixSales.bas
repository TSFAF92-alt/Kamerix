Attribute VB_Name = "modKamerixSales"
Option Explicit
Option Private Module

'==============================================================================
' KAMERIX - Service de vente comptant
'------------------------------------------------------------------------------
' Cycle       : StartCashSale (voiture -> VENTE_EN_COURS, mouvement EN_COURS,
'               dossier de CV A_DEPOSER) ; RecordSalePayment (recu REC-..., au
'               dernier paiement : mouvement SOLDE et facture FAC-... emise) ;
'               SubmitCv (vente soldee exigee) ; SignCv (voiture -> VENDU).
' Annulation  : CancelSale, tant que le CV n'est pas signe : remboursement RMB-...
'               des sommes versees (frais du CV deduits si l'acheteur se desiste),
'               facture annulee si elle existe, voiture remise en stock.
' Droits      : Ventes (ecriture) pour la vente et le CV ; Paiements (ecriture)
'               pour un paiement ; VENTE_ANNULER pour l'annulation, et REMBOURSER
'               s'il y a quelque chose a rembourser.
' Marge       : prix - cout de revient (achat + commission d'achat + charges)
'               - commission de vente - frais du CV. Une vente sous le cout est
'               acceptee : EstimatedMargin permet a l'ecran d'avertir avant.
' Transaction : chaque operation est tout ou rien ; toutes les regles et donnees
'               sont controlees avant la premiere ecriture.
' Dependances : DataAccess, Car, Party, Movement, CashSale (IContract), Payment,
'               Invoice, CvTracking, modKamerixPurchase, modKamerixNumbering,
'               modKamerixSettings, modKamerixReference, modKamerixSession.
'==============================================================================

' Colonnes renvoyees par SaleRows (tableau colonne x ligne pour ListBox.Column)
Public Const SALE_COL_ID As Long = 0
Public Const SALE_COL_DATE As Long = 1
Public Const SALE_COL_CLIENT As Long = 2
Public Const SALE_COL_CAR As Long = 3
Public Const SALE_COL_PRICE As Long = 4
Public Const SALE_COL_PAID As Long = 5
Public Const SALE_COL_BALANCE As Long = 6
Public Const SALE_COL_STATUS As Long = 7
Public Const SALE_COL_CV As Long = 8
Public Const SALE_COL_STATUS_CODE As Long = 9
Public Const SALE_COL_CV_CODE As Long = 10
Private Const SALE_COL_COUNT As Long = 11

'==============================================================================
' VENTE
'==============================================================================

' Demarre une vente comptant directe (voiture en stock) et retourne l'identifiant du mouvement.
Public Function StartCashSale(ByVal pCarId As Long, ByVal pBuyerId As Long, ByVal pPrice As Currency, _
                              Optional ByVal pCommission As Currency = 0, Optional ByVal pIntermediaryId As Long = 0, _
                              Optional ByVal pPlaceId As Long = 0, Optional ByVal pDate As Variant) As Long
    Const SRC As String = "modKamerixSales.StartCashSale"
    Dim objData As DataAccess
    Dim objCar As Car
    Dim objBuyer As Party
    Dim objSale As Movement
    Dim objCv As CvTracking
    Dim blnOwn As Boolean
    Dim lngErr As Long
    Dim strErr As String
    Dim strSrc As String

    RequireWrite MOD_SALES
    Set objData = New DataAccess

    Set objCar = objData.LoadCar(pCarId)
    If objCar Is Nothing Then Err.Raise ERR_SALE, SRC, "Voiture introuvable (identifiant " & pCarId & ")."
    If objCar.Status <> CAR_STATUS_IN_STOCK Then
        Err.Raise ERR_SALE, SRC, "Seule une voiture en stock peut " & ChrW(234) & "tre vendue directement (statut : " & objCar.Status & ")."
    End If
    If objData.LoadPurchaseOfCar(pCarId) Is Nothing Then
        Err.Raise ERR_SALE, SRC, "L'achat de cette voiture n'est pas enregistr" & ChrW(233) & " : enregistrez-le avant la vente."
    End If

    Set objBuyer = objData.LoadParty(pBuyerId)
    If objBuyer Is Nothing Then Err.Raise ERR_SALE, SRC, "Acheteur introuvable (identifiant " & pBuyerId & ")."
    If Not objBuyer.IsBuyer Then
        Err.Raise ERR_SALE, SRC, objBuyer.Name & " n'est pas enregistr" & ChrW(233) & " comme acheteur."
    End If
    If pIntermediaryId <> 0 Then
        If objData.LoadIntermediary(pIntermediaryId) Is Nothing Then
            Err.Raise ERR_SALE, SRC, "Interm" & ChrW(233) & "diaire introuvable (identifiant " & pIntermediaryId & ")."
        End If
    End If
    If pPlaceId <> 0 Then
        If Len(PlaceTypeOf(pPlaceId)) = 0 Then Err.Raise ERR_SALE, SRC, "Lieu introuvable (identifiant " & pPlaceId & ")."
    End If

    Set objSale = New Movement
    objSale.MovementType = MOVEMENT_TYPE_SALE
    objSale.SettlementMode = SETTLEMENT_CASH
    objSale.CarId = pCarId
    objSale.PartyId = pBuyerId
    objSale.Price = pPrice
    objSale.Commission = pCommission
    objSale.IntermediaryId = pIntermediaryId
    objSale.PlaceId = pPlaceId
    If Not IsMissing(pDate) Then objSale.MovementDate = CDate(pDate)
    objSale.SetUserId CurrentUserId()
    objSale.Validate

    objCar.StartDirectSale

    blnOwn = BeginIfNeeded()
    On Error GoTo EH
    objData.SaveMovement objSale
    objData.SaveCar objCar
    Set objCv = New CvTracking
    objCv.MovementId = objSale.Id
    objData.SaveCvTracking objCv
    CommitIfOwn blnOwn

    StartCashSale = objSale.Id
    Exit Function

EH:
    lngErr = Err.Number: strErr = Err.Description: strSrc = Err.Source
    RollbackIfOwn blnOwn
    Err.Raise lngErr, strSrc, strErr
End Function

' Contrat de la vente (vu par l'interface IContract) : prix, encaisse, solde.
' Lecture : module Ventes.
Public Function LoadSaleContract(ByVal pMovementId As Long) As IContract
    Dim objData As DataAccess
    Dim objSale As Movement
    Dim objContract As CashSale

    RequireRead MOD_SALES
    Set objData = New DataAccess
    Set objSale = LoadSaleMovement(objData, pMovementId, "modKamerixSales.LoadSaleContract")

    Set objContract = New CashSale
    objContract.Init objSale, NetPaid(objData, pMovementId)
    Set LoadSaleContract = objContract
End Function

'==============================================================================
' PAIEMENTS ET FACTURE
'==============================================================================

' Enregistre un paiement de la vente et retourne le numero du recu (REC-...).
' Au dernier paiement, la vente est soldee et la facture emise.
Public Function RecordSalePayment(ByVal pMovementId As Long, ByVal pAmount As Currency, ByVal pPaymentMode As String, _
                                  Optional ByVal pDate As Variant) As String
    Const SRC As String = "modKamerixSales.RecordSalePayment"
    Dim objData As DataAccess
    Dim objSale As Movement
    Dim objContract As IContract
    Dim objCash As CashSale
    Dim objPayment As Payment
    Dim objInvoice As Invoice
    Dim dtmPayment As Date
    Dim blnOwn As Boolean
    Dim lngErr As Long
    Dim strErr As String
    Dim strSrc As String

    RequireWrite MOD_PAYMENTS
    Set objData = New DataAccess
    Set objSale = LoadSaleMovement(objData, pMovementId, SRC)

    Set objCash = New CashSale
    objCash.Init objSale, NetPaid(objData, pMovementId)
    Set objContract = objCash

    If objSale.Status <> MOVEMENT_STATUS_IN_PROGRESS Then
        Err.Raise ERR_SALE, SRC, "Cette vente n'attend plus de paiement (statut : " & objSale.Status & ")."
    End If
    If Not objContract.CanReceive(pAmount) Then
        Err.Raise ERR_SALE, SRC, "Montant refus" & ChrW(233) & " : il doit " & ChrW(234) & "tre sup" & ChrW(233) & "rieur " & _
            ChrW(224) & " z" & ChrW(233) & "ro et ne pas d" & ChrW(233) & "passer le solde (" & Format$(objContract.Balance, "#,##0") & ")."
    End If

    If IsMissing(pDate) Then dtmPayment = Date Else dtmPayment = CDate(pDate)
    Set objPayment = New Payment
    objPayment.MovementId = pMovementId
    objPayment.Direction = PAYMENT_DIRECTION_IN
    objPayment.Amount = pAmount
    objPayment.PaymentMode = pPaymentMode
    objPayment.PaymentDate = dtmPayment
    objPayment.SetUserId CurrentUserId()
    objPayment.SetReceiptNumber "REC-provisoire"     ' remplace par le vrai numero dans la transaction
    objPayment.Validate

    blnOwn = BeginIfNeeded()
    On Error GoTo EH
    objPayment.SetReceiptNumber NextDocumentNumber(DOC_RECEIPT, dtmPayment)
    objData.SavePayment objPayment

    ' Dernier paiement : vente soldee et facture emise pour le prix total
    If objContract.Balance - pAmount <= 0 Then
        objSale.MarkSettled
        objData.SaveMovement objSale

        Set objInvoice = New Invoice
        objInvoice.Prepare NextDocumentNumber(DOC_INVOICE, dtmPayment), pMovementId, objSale.PartyId, dtmPayment, _
                           CarDesignation(objData, objSale.CarId), objSale.Price, TaxRate(), CurrentUserId()
        objData.SaveInvoice objInvoice
        objData.LinkPaymentsToInvoice pMovementId, objInvoice.Id
    End If
    CommitIfOwn blnOwn

    RecordSalePayment = objPayment.ReceiptNumber
    Exit Function

EH:
    lngErr = Err.Number: strErr = Err.Description: strSrc = Err.Source
    RollbackIfOwn blnOwn
    Err.Raise lngErr, strSrc, strErr
End Function

' Facture de la vente (Nothing tant qu'elle n'est pas soldee). Lecture : module Factures.
Public Function InvoiceOfSale(ByVal pMovementId As Long) As Invoice
    Dim objData As DataAccess

    RequireRead MOD_INVOICES
    Set objData = New DataAccess
    Set InvoiceOfSale = objData.LoadInvoiceOfMovement(pMovementId)
End Function

'==============================================================================
' CERTIFICAT DE VENTE
'==============================================================================

' Suivi du CV de la vente. Lecture : module Ventes.
Public Function CvOfSale(ByVal pMovementId As Long) As CvTracking
    Dim objData As DataAccess

    RequireRead MOD_SALES
    Set objData = New DataAccess
    Set CvOfSale = objData.LoadCvTrackingOfMovement(pMovementId)
End Function

' Depot du dossier de CV au commissariat (vente soldee exigee), avec les frais payes
Public Sub SubmitCv(ByVal pMovementId As Long, ByVal pPoliceStationId As Long, ByVal pFilingDate As Date, _
                    ByVal pFees As Currency)
    Const SRC As String = "modKamerixSales.SubmitCv"
    Dim objData As DataAccess
    Dim objSale As Movement
    Dim objCv As CvTracking

    RequireWrite MOD_SALES
    Set objData = New DataAccess
    Set objSale = LoadSaleMovement(objData, pMovementId, SRC)

    If objSale.Status <> MOVEMENT_STATUS_SETTLED Then
        Err.Raise ERR_CV, SRC, "Le dossier de CV ne peut " & ChrW(234) & "tre d" & ChrW(233) & "pos" & ChrW(233) & _
            " qu'une fois la vente sold" & ChrW(233) & "e."
    End If
    If PlaceTypeOf(pPoliceStationId) <> PLACE_TYPE_POLICE_STATION Then
        Err.Raise ERR_CV, SRC, "Le lieu choisi n'est pas un commissariat (identifiant " & pPoliceStationId & ")."
    End If

    Set objCv = RequireCv(objData, pMovementId, SRC)
    objCv.Submit pPoliceStationId, pFilingDate, pFees
    objData.SaveCvTracking objCv
End Sub

' CV signe : la voiture passe a VENDU (definitif)
Public Sub SignCv(ByVal pMovementId As Long, ByVal pCvNumber As String, ByVal pSignatureDate As Date)
    Const SRC As String = "modKamerixSales.SignCv"
    Dim objData As DataAccess
    Dim objSale As Movement
    Dim objCv As CvTracking
    Dim objCar As Car
    Dim blnOwn As Boolean
    Dim lngErr As Long
    Dim strErr As String
    Dim strSrc As String

    RequireWrite MOD_SALES
    Set objData = New DataAccess
    Set objSale = LoadSaleMovement(objData, pMovementId, SRC)
    Set objCv = RequireCv(objData, pMovementId, SRC)
    Set objCar = objData.LoadCar(objSale.CarId)

    objCv.MarkSigned pCvNumber, pSignatureDate
    objCar.MarkAsSold

    blnOwn = BeginIfNeeded()
    On Error GoTo EH
    objData.SaveCvTracking objCv
    objData.SaveCar objCar
    CommitIfOwn blnOwn
    Exit Sub

EH:
    lngErr = Err.Number: strErr = Err.Description: strSrc = Err.Source
    RollbackIfOwn blnOwn
    Err.Raise lngErr, strSrc, strErr
End Sub

'==============================================================================
' ANNULATION
'==============================================================================

' Montant qui serait rembourse si la vente etait annulee maintenant :
' sommes versees, moins les frais du CV deja payes si l'acheteur se desiste (jamais negatif).
Public Function RefundAmountIfCancelled(ByVal pMovementId As Long, ByVal pBuyerWithdrawal As Boolean) As Currency
    Dim objData As DataAccess
    Dim curRefund As Currency

    RequireRead MOD_SALES
    Set objData = New DataAccess
    curRefund = NetPaid(objData, pMovementId)
    If pBuyerWithdrawal Then curRefund = curRefund - CvFees(objData, pMovementId)
    If curRefund < 0 Then curRefund = 0
    RefundAmountIfCancelled = curRefund
End Function

' Annule la vente (CV non signe). Retourne le numero du remboursement (RMB-...),
' ou une chaine vide s'il n'y a rien a rembourser.
'   pBuyerWithdrawal : Vrai si l'acheteur se desiste (frais du CV a sa charge)
'   pRefundMode      : mode du remboursement (obligatoire s'il y a un remboursement)
Public Function CancelSale(ByVal pMovementId As Long, ByVal pReason As String, ByVal pBuyerWithdrawal As Boolean, _
                           Optional ByVal pRefundMode As String = "") As String
    Const SRC As String = "modKamerixSales.CancelSale"
    Dim objData As DataAccess
    Dim objSale As Movement
    Dim objCar As Car
    Dim objCv As CvTracking
    Dim objInvoice As Invoice
    Dim objRefund As Payment
    Dim curRefund As Currency
    Dim blnOwn As Boolean
    Dim lngErr As Long
    Dim strErr As String
    Dim strSrc As String

    RequireRight RIGHT_CANCEL_SALE
    Set objData = New DataAccess
    Set objSale = LoadSaleMovement(objData, pMovementId, SRC)

    If objSale.Status = MOVEMENT_STATUS_CANCELLED Then Err.Raise ERR_SALE, SRC, "Cette vente est d" & ChrW(233) & "j" & ChrW(224) & " annul" & ChrW(233) & "e."
    Set objCv = objData.LoadCvTrackingOfMovement(pMovementId)
    If Not objCv Is Nothing Then
        If objCv.Status = CV_STATUS_SIGNED Then
            Err.Raise ERR_SALE, SRC, "Le CV est sign" & ChrW(233) & " : la vente ne peut plus " & ChrW(234) & "tre annul" & ChrW(233) & "e."
        End If
    End If

    curRefund = NetPaid(objData, pMovementId)
    If pBuyerWithdrawal And Not objCv Is Nothing Then curRefund = curRefund - objCv.FeesPaid
    If curRefund < 0 Then curRefund = 0

    If curRefund > 0 Then
        RequireRight RIGHT_REFUND
        Set objRefund = New Payment
        objRefund.MovementId = pMovementId
        objRefund.Direction = PAYMENT_DIRECTION_REFUND
        objRefund.Amount = curRefund
        objRefund.PaymentMode = pRefundMode
        objRefund.SetUserId CurrentUserId()
        objRefund.SetReceiptNumber "RMB-provisoire"
        objRefund.Validate
    End If

    objSale.Cancel pReason
    objSale.Validate
    Set objCar = objData.LoadCar(objSale.CarId)
    objCar.CancelSale pReason
    Set objInvoice = objData.LoadInvoiceOfMovement(pMovementId)
    If Not objInvoice Is Nothing Then
        If objInvoice.Status = INVOICE_STATUS_ISSUED Then objInvoice.Cancel Else Set objInvoice = Nothing
    End If

    blnOwn = BeginIfNeeded()
    On Error GoTo EH
    If Not objRefund Is Nothing Then
        objRefund.SetReceiptNumber NextDocumentNumber(DOC_REFUND)
        objData.SavePayment objRefund
    End If
    objData.SaveMovement objSale
    objData.SaveCar objCar
    If Not objInvoice Is Nothing Then objData.SaveInvoice objInvoice
    CommitIfOwn blnOwn

    If Not objRefund Is Nothing Then CancelSale = objRefund.ReceiptNumber
    Exit Function

EH:
    lngErr = Err.Number: strErr = Err.Description: strSrc = Err.Source
    RollbackIfOwn blnOwn
    Err.Raise lngErr, strSrc, strErr
End Function

'==============================================================================
' MARGE
'==============================================================================

' Marge prevue avant la vente : prix - cout de revient - commission de vente
' (les frais du CV ne sont pas encore connus). Negative : l'ecran avertit.
Public Function EstimatedMargin(ByVal pCarId As Long, ByVal pPrice As Currency, Optional ByVal pCommission As Currency = 0) As Currency
    EstimatedMargin = pPrice - CarCostPrice(pCarId) - pCommission
End Function

' Marge d'une vente : prix - cout de revient - commission de vente - frais du CV
Public Function SaleMargin(ByVal pMovementId As Long) As Currency
    Dim objData As DataAccess
    Dim objSale As Movement

    RequireRead MOD_SALES
    Set objData = New DataAccess
    Set objSale = LoadSaleMovement(objData, pMovementId, "modKamerixSales.SaleMargin")
    SaleMargin = objSale.Price - CarCostPrice(objSale.CarId) - objSale.Commission - CvFees(objData, pMovementId)
End Function

'==============================================================================
' ECRANS (lot 5.4c) : listes et libelles
'==============================================================================

' Ventes comptant, les plus recentes d'abord, pretes pour ListBox.Column.
' pStatus : code de statut du mouvement, vide = toutes. Empty si aucune vente.
' Montants formates ; encaisse = encaissements - remboursements ; solde vide si annulee.
' Lecture : module Ventes.
Public Function SaleRows(Optional ByVal pStatus As String = "") As Variant
    Dim AdoRs As ADODB.Recordset
    Dim strSql As String
    Dim vData() As Variant
    Dim lngRows As Long
    Dim i As Long
    Dim curPaid As Currency
    Dim strStatus As String
    Dim strCv As String

    RequireRead MOD_SALES

    strSql = "SELECT mv.[MouvementID], mv.[DateMouvement], t.[Nom] AS Client, " & _
             "b.[Nom] & ' ' & m.[Nom] & ' - ' & v.[Chassis] AS Voiture, mv.[Prix], mv.[Statut], cv.[Statut] AS StatutCV, " & _
             "(SELECT SUM(p.[Montant]) FROM [Paiement] AS p WHERE p.[MouvementID] = mv.[MouvementID] AND p.[Sens] = ?) AS Encaisse, " & _
             "(SELECT SUM(r.[Montant]) FROM [Paiement] AS r WHERE r.[MouvementID] = mv.[MouvementID] AND r.[Sens] = ?) AS Rembourse " & _
             "FROM (((([Mouvement] AS mv INNER JOIN [Tiers] AS t ON mv.[TiersID] = t.[TiersID]) " & _
             "INNER JOIN [Voiture] AS v ON mv.[VoitureID] = v.[VoitureID]) " & _
             "INNER JOIN [Modele] AS m ON v.[ModeleID] = m.[ModeleID]) " & _
             "INNER JOIN [Marque] AS b ON m.[MarqueID] = b.[MarqueID]) " & _
             "LEFT JOIN [SuiviCV] AS cv ON cv.[MouvementID] = mv.[MouvementID] " & _
             "WHERE mv.[TypeMouvement] = ? AND mv.[ModeReglement] = ? "

    If Len(pStatus) = 0 Then
        Set AdoRs = QueryRecordset(strSql & "ORDER BY mv.[DateMouvement] DESC, mv.[MouvementID] DESC", _
                                   PAYMENT_DIRECTION_IN, PAYMENT_DIRECTION_REFUND, MOVEMENT_TYPE_SALE, SETTLEMENT_CASH)
    Else
        Set AdoRs = QueryRecordset(strSql & "AND mv.[Statut] = ? ORDER BY mv.[DateMouvement] DESC, mv.[MouvementID] DESC", _
                                   PAYMENT_DIRECTION_IN, PAYMENT_DIRECTION_REFUND, MOVEMENT_TYPE_SALE, SETTLEMENT_CASH, pStatus)
    End If

    Do While Not AdoRs.EOF
        lngRows = lngRows + 1
        AdoRs.MoveNext
    Loop
    If lngRows = 0 Then
        AdoRs.Close
        Exit Function
    End If

    ReDim vData(0 To SALE_COL_COUNT - 1, 0 To lngRows - 1)
    AdoRs.MoveFirst
    Do While Not AdoRs.EOF
        curPaid = CurrencyOrZero(AdoRs.Fields("Encaisse").Value) - CurrencyOrZero(AdoRs.Fields("Rembourse").Value)
        strStatus = CStr(AdoRs.Fields("Statut").Value)
        strCv = TextOfValue(AdoRs.Fields("StatutCV").Value)

        vData(SALE_COL_ID, i) = CLng(AdoRs.Fields("MouvementID").Value)
        vData(SALE_COL_DATE, i) = Format$(AdoRs.Fields("DateMouvement").Value, "dd/mm/yyyy")
        vData(SALE_COL_CLIENT, i) = CStr(AdoRs.Fields("Client").Value)
        vData(SALE_COL_CAR, i) = CStr(AdoRs.Fields("Voiture").Value)
        vData(SALE_COL_PRICE, i) = Format$(AdoRs.Fields("Prix").Value, "#,##0")
        vData(SALE_COL_PAID, i) = Format$(curPaid, "#,##0")
        If strStatus = MOVEMENT_STATUS_CANCELLED Then
            vData(SALE_COL_BALANCE, i) = ""
        Else
            vData(SALE_COL_BALANCE, i) = Format$(CCur(AdoRs.Fields("Prix").Value) - curPaid, "#,##0")
        End If
        vData(SALE_COL_STATUS, i) = SaleStatusLabel(strStatus)
        vData(SALE_COL_CV, i) = CvStatusLabel(strCv)
        vData(SALE_COL_STATUS_CODE, i) = strStatus
        vData(SALE_COL_CV_CODE, i) = strCv
        i = i + 1
        AdoRs.MoveNext
    Loop
    AdoRs.Close

    SaleRows = vData
End Function

' Voitures qui peuvent etre vendues : en stock et achetees. Colonnes : VoitureID, Libelle.
' Jointure sur le mouvement d'achat non annule (une voiture n'en a qu'un : regle de
' modKamerixPurchase), plutot qu'une sous-requete EXISTS parametree.
' Lecture : module Ventes.
Public Function SaleableCarList() As ADODB.Recordset
    RequireRead MOD_SALES
    Set SaleableCarList = QueryRecordset( _
        "SELECT v.[VoitureID], b.[Nom] & ' ' & m.[Nom] & ' - ' & v.[Chassis] AS Libelle " & _
        "FROM (([Voiture] AS v INNER JOIN [Modele] AS m ON v.[ModeleID] = m.[ModeleID]) " & _
        "INNER JOIN [Marque] AS b ON m.[MarqueID] = b.[MarqueID]) " & _
        "INNER JOIN [Mouvement] AS a ON a.[VoitureID] = v.[VoitureID] " & _
        "WHERE v.[Statut] = ? AND a.[TypeMouvement] = ? AND a.[Statut] <> ? " & _
        "ORDER BY b.[Nom], m.[Nom], v.[Chassis]", _
        CAR_STATUS_IN_STOCK, MOVEMENT_TYPE_PURCHASE, MOVEMENT_STATUS_CANCELLED)
End Function

' Libelle d'une vente : "Client - Toyota Corolla - CHASSIS" (vide si introuvable). Lecture : module Ventes.
Public Function SaleLabel(ByVal pMovementId As Long) As String
    Dim vLabel As Variant

    RequireRead MOD_SALES
    vLabel = QueryValue( _
        "SELECT t.[Nom] & ' - ' & b.[Nom] & ' ' & m.[Nom] & ' - ' & v.[Chassis] " & _
        "FROM ((([Mouvement] AS mv INNER JOIN [Tiers] AS t ON mv.[TiersID] = t.[TiersID]) " & _
        "INNER JOIN [Voiture] AS v ON mv.[VoitureID] = v.[VoitureID]) " & _
        "INNER JOIN [Modele] AS m ON v.[ModeleID] = m.[ModeleID]) " & _
        "INNER JOIN [Marque] AS b ON m.[MarqueID] = b.[MarqueID] WHERE mv.[MouvementID] = ?", pMovementId)
    If Not IsNull(vLabel) Then SaleLabel = CStr(vLabel)
End Function

Public Function SaleStatusLabel(ByVal pStatus As String) As String
    Select Case pStatus
        Case MOVEMENT_STATUS_IN_PROGRESS: SaleStatusLabel = "En cours"
        Case MOVEMENT_STATUS_SETTLED:     SaleStatusLabel = "Sold" & ChrW(233) & "e"
        Case MOVEMENT_STATUS_CANCELLED:   SaleStatusLabel = "Annul" & ChrW(233) & "e"
        Case Else:                        SaleStatusLabel = pStatus
    End Select
End Function

Public Function CvStatusLabel(ByVal pStatus As String) As String
    Select Case pStatus
        Case CV_STATUS_TO_FILE:     CvStatusLabel = ChrW(192) & " d" & ChrW(233) & "poser"
        Case CV_STATUS_IN_PROGRESS: CvStatusLabel = "D" & ChrW(233) & "pos" & ChrW(233)
        Case CV_STATUS_SIGNED:      CvStatusLabel = "Sign" & ChrW(233)
        Case Else:                  CvStatusLabel = pStatus
    End Select
End Function

'==============================================================================
' PROCEDURES INTERNES
'==============================================================================

' Mouvement de vente comptant ; erreur s'il n'existe pas ou n'en est pas un
Private Function LoadSaleMovement(ByVal pData As DataAccess, ByVal pMovementId As Long, ByVal pSource As String) As Movement
    Dim objSale As Movement

    Set objSale = pData.LoadMovement(pMovementId)
    If objSale Is Nothing Then Err.Raise ERR_SALE, pSource, "Vente introuvable (identifiant " & pMovementId & ")."
    If objSale.MovementType <> MOVEMENT_TYPE_SALE Or objSale.SettlementMode <> SETTLEMENT_CASH Then
        Err.Raise ERR_SALE, pSource, "Le mouvement " & pMovementId & " n'est pas une vente comptant."
    End If
    Set LoadSaleMovement = objSale
End Function

Private Function RequireCv(ByVal pData As DataAccess, ByVal pMovementId As Long, ByVal pSource As String) As CvTracking
    Dim objCv As CvTracking

    Set objCv = pData.LoadCvTrackingOfMovement(pMovementId)
    If objCv Is Nothing Then Err.Raise ERR_CV, pSource, "Aucun suivi de CV pour la vente " & pMovementId & "."
    Set RequireCv = objCv
End Function

' Encaisse net : encaissements - remboursements
Private Function NetPaid(ByVal pData As DataAccess, ByVal pMovementId As Long) As Currency
    NetPaid = pData.PaymentsTotal(pMovementId, PAYMENT_DIRECTION_IN) - pData.PaymentsTotal(pMovementId, PAYMENT_DIRECTION_REFUND)
End Function

Private Function CvFees(ByVal pData As DataAccess, ByVal pMovementId As Long) As Currency
    Dim objCv As CvTracking

    Set objCv = pData.LoadCvTrackingOfMovement(pMovementId)
    If Not objCv Is Nothing Then CvFees = objCv.FeesPaid
End Function

' Designation figee sur la facture : "Toyota Corolla 2015 - chassis X - immatriculation Y"
Private Function CarDesignation(ByVal pData As DataAccess, ByVal pCarId As Long) As String
    Dim objCar As Car
    Dim strText As String

    Set objCar = pData.LoadCar(pCarId)
    strText = "V" & ChrW(233) & "hicule " & ModelLabel(objCar.ModelId)
    If Not IsNull(objCar.ModelYear) Then strText = strText & " " & objCar.ModelYear
    strText = strText & " - ch" & ChrW(226) & "ssis " & objCar.Chassis
    If Len(objCar.PlateNumber) > 0 Then strText = strText & " - immatriculation " & objCar.PlateNumber
    If Len(objCar.Color) > 0 Then strText = strText & " - " & objCar.Color
    CarDesignation = strText
End Function

' Transaction : ouverte ici seulement si l'appelant n'en a pas deja une
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

Private Function CurrencyOrZero(ByVal pValue As Variant) As Currency
    If Not IsNull(pValue) And Not IsEmpty(pValue) Then CurrencyOrZero = CCur(pValue)
End Function
