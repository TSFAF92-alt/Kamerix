VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Classe d'accès aux données : lit et écrit Car, Party, Intermediary, Movement, Payment,
' Invoice et CvTracking dans la base.
' Toutes les requêtes passent par modKamerixSql (paramètres "?", transactions).
' Cette classe n'ouvre pas de transaction elle-même : si elle est appelée entre
' BeginTransaction et CommitTransaction, ses écritures en font partie.
' Les objets sont validés (Validate) avant chaque enregistrement.

' Colonnes lues pour un mouvement (achat ou vente)
Private Const SQL_MOVEMENT As String = _
    "SELECT [MouvementID], [TypeMouvement], [VoitureID], [TiersID], [IntermediaireID], [LieuID], [DateMouvement], " & _
    "[Prix], [Commission], [ModeReglement], [Statut], [TypeRetenue], [ValeurRetenue], [DateLimiteSolde], " & _
    "[DateAnnulation], [MotifAnnulation], [UtilisateurID] FROM [Mouvement] "

' ---------- Voiture ----------

' Retourne la voiture, ou Nothing si l'identifiant n'existe pas
Public Function LoadCar(ByVal pId As Long) As Car
    Dim AdoRs As ADODB.Recordset
    Dim v As Car

    Set AdoRs = QueryRecordset( _
        "SELECT [VoitureID], [ModeleID], [Annee], [Couleur], [Chassis], [Immatriculation], [Kilometrage], " & _
        "[Statut], [MotifRetrait] FROM [Voiture] WHERE [VoitureID] = ?", pId)

    If Not AdoRs.EOF Then
        Set v = New Car
        v.Hydrate CLng(AdoRs.Fields("VoitureID").Value), CStr(AdoRs.Fields("Statut").Value), _
                   TextOf(AdoRs.Fields("MotifRetrait").Value)
        v.HydrateData CLng(AdoRs.Fields("ModeleID").Value), AdoRs.Fields("Annee").Value, _
                          AdoRs.Fields("Couleur").Value, AdoRs.Fields("Chassis").Value, _
                          AdoRs.Fields("Immatriculation").Value, AdoRs.Fields("Kilometrage").Value
    End If
    AdoRs.Close

    Set LoadCar = v
End Function

' Insère la voiture (identifiant 0) ou met à jour la voiture existante, statut compris
Public Sub SaveCar(ByVal pCar As Car)
    Dim lngId As Long
    Dim lngRows As Long

    pCar.Validate
    RequireUniqueChassis pCar

    If pCar.Id = 0 Then
        lngId = ExecuteInsert( _
            "INSERT INTO [Voiture] ([ModeleID], [Annee], [Couleur], [Chassis], [Immatriculation], [Kilometrage], " & _
            "[Statut], [MotifRetrait]) VALUES (?, ?, ?, ?, ?, ?, ?, ?)", _
            pCar.ModelId, pCar.ModelYear, NullIfEmpty(pCar.Color), pCar.Chassis, _
            NullIfEmpty(pCar.PlateNumber), pCar.Mileage, pCar.Status, _
            NullIfEmpty(pCar.WithdrawalReason))
        pCar.SetId lngId
    Else
        lngRows = ExecuteNonQuery( _
            "UPDATE [Voiture] SET [ModeleID] = ?, [Annee] = ?, [Couleur] = ?, [Chassis] = ?, " & _
            "[Immatriculation] = ?, [Kilometrage] = ?, [Statut] = ?, [MotifRetrait] = ? WHERE [VoitureID] = ?", _
            pCar.ModelId, pCar.ModelYear, NullIfEmpty(pCar.Color), pCar.Chassis, _
            NullIfEmpty(pCar.PlateNumber), pCar.Mileage, pCar.Status, _
            NullIfEmpty(pCar.WithdrawalReason), pCar.Id)
        If lngRows = 0 Then
            Err.Raise ERR_DATA_NOT_FOUND, "DataAccess.SaveCar", _
                "Voiture introuvable (identifiant " & pCar.Id & ")."
        End If
    End If
End Sub

' ---------- Tiers ----------

' Retourne le tiers, ou Nothing si l'identifiant n'existe pas
Public Function LoadParty(ByVal pId As Long) As Party
    Dim AdoRs As ADODB.Recordset
    Dim t As Party

    Set AdoRs = QueryRecordset( _
        "SELECT [TiersID], [Nom], [Telephone], [TypePiece], [NumeroPiece], [Adresse], [EstAcheteur], [EstVendeur] " & _
        "FROM [Tiers] WHERE [TiersID] = ?", pId)

    If Not AdoRs.EOF Then
        Set t = New Party
        t.Hydrate CLng(AdoRs.Fields("TiersID").Value), AdoRs.Fields("Nom").Value, AdoRs.Fields("Telephone").Value, _
                   AdoRs.Fields("TypePiece").Value, AdoRs.Fields("NumeroPiece").Value, AdoRs.Fields("Adresse").Value, _
                   AdoRs.Fields("EstAcheteur").Value, AdoRs.Fields("EstVendeur").Value
    End If
    AdoRs.Close

    Set LoadParty = t
End Function

' Insère le tiers (identifiant 0) ou met à jour le tiers existant
Public Sub SaveParty(ByVal pParty As Party)
    Dim lngId As Long
    Dim lngRows As Long

    pParty.Validate

    If pParty.Id = 0 Then
        lngId = ExecuteInsert( _
            "INSERT INTO [Tiers] ([Nom], [Telephone], [TypePiece], [NumeroPiece], [Adresse], [EstAcheteur], [EstVendeur]) " & _
            "VALUES (?, ?, ?, ?, ?, ?, ?)", _
            pParty.Name, NullIfEmpty(pParty.Phone), NullIfEmpty(pParty.IdType), _
            NullIfEmpty(pParty.IdNumber), NullIfEmpty(pParty.Address), _
            CBool(pParty.IsBuyer), CBool(pParty.IsSeller))
        pParty.SetId lngId
    Else
        lngRows = ExecuteNonQuery( _
            "UPDATE [Tiers] SET [Nom] = ?, [Telephone] = ?, [TypePiece] = ?, [NumeroPiece] = ?, " & _
            "[Adresse] = ?, [EstAcheteur] = ?, [EstVendeur] = ? WHERE [TiersID] = ?", _
            pParty.Name, NullIfEmpty(pParty.Phone), NullIfEmpty(pParty.IdType), _
            NullIfEmpty(pParty.IdNumber), NullIfEmpty(pParty.Address), _
            CBool(pParty.IsBuyer), CBool(pParty.IsSeller), pParty.Id)
        If lngRows = 0 Then
            Err.Raise ERR_DATA_NOT_FOUND, "DataAccess.SaveParty", _
                "Tiers introuvable (identifiant " & pParty.Id & ")."
        End If
    End If
End Sub

' ---------- Intermédiaire ----------

' Retourne l'intermédiaire, ou Nothing si l'identifiant n'existe pas
Public Function LoadIntermediary(ByVal pId As Long) As Intermediary
    Dim AdoRs As ADODB.Recordset
    Dim i As Intermediary

    Set AdoRs = QueryRecordset( _
        "SELECT [IntermediaireID], [Nom], [Telephone], [VilleID] FROM [Intermediaire] WHERE [IntermediaireID] = ?", pId)
    If Not AdoRs.EOF Then
        Set i = New Intermediary
        i.Hydrate CLng(AdoRs.Fields("IntermediaireID").Value), AdoRs.Fields("Nom").Value, _
                  AdoRs.Fields("Telephone").Value, CLng(AdoRs.Fields("VilleID").Value)
    End If
    AdoRs.Close
    Set LoadIntermediary = i
End Function

' Insère l'intermédiaire (identifiant 0) ou met à jour l'intermédiaire existant.
' La ville doit être un lieu de type VILLE.
Public Sub SaveIntermediary(ByVal pIntermediary As Intermediary)
    Dim lngId As Long
    Dim lngRows As Long

    pIntermediary.Validate
    If PlaceTypeOf(pIntermediary.CityId) <> PLACE_TYPE_CITY Then
        Err.Raise ERR_INTERMEDIARY_INVALID, "DataAccess.SaveIntermediary", _
            "La ville de l'intermédiaire est introuvable ou n'est pas une ville (identifiant " & pIntermediary.CityId & ")."
    End If

    If pIntermediary.Id = 0 Then
        lngId = ExecuteInsert( _
            "INSERT INTO [Intermediaire] ([Nom], [Telephone], [VilleID]) VALUES (?, ?, ?)", _
            pIntermediary.Name, NullIfEmpty(pIntermediary.Phone), pIntermediary.CityId)
        pIntermediary.SetId lngId
    Else
        lngRows = ExecuteNonQuery( _
            "UPDATE [Intermediaire] SET [Nom] = ?, [Telephone] = ?, [VilleID] = ? WHERE [IntermediaireID] = ?", _
            pIntermediary.Name, NullIfEmpty(pIntermediary.Phone), pIntermediary.CityId, pIntermediary.Id)
        If lngRows = 0 Then
            Err.Raise ERR_DATA_NOT_FOUND, "DataAccess.SaveIntermediary", _
                "Intermédiaire introuvable (identifiant " & pIntermediary.Id & ")."
        End If
    End If
End Sub

' ---------- Mouvement (achat ou vente) ----------

' Retourne le mouvement, ou Nothing si l'identifiant n'existe pas
Public Function LoadMovement(ByVal pId As Long) As Movement
    Set LoadMovement = MovementFromQuery(SQL_MOVEMENT & "WHERE [MouvementID] = ?", pId)
End Function

' Achat d'une voiture (Nothing si la voiture n'a pas d'achat ; un achat annulé est ignoré)
Public Function LoadPurchaseOfCar(ByVal pCarId As Long) As Movement
    Set LoadPurchaseOfCar = MovementFromQuery(SQL_MOVEMENT & _
        "WHERE [VoitureID] = ? AND [TypeMouvement] = ? AND [Statut] <> ?", _
        pCarId, MOVEMENT_TYPE_PURCHASE, MOVEMENT_STATUS_CANCELLED)
End Function

' Insère le mouvement (identifiant 0) ou met à jour le mouvement existant
Public Sub SaveMovement(ByVal pMovement As Movement)
    Dim lngId As Long
    Dim lngRows As Long

    pMovement.Validate

    If pMovement.Id = 0 Then
        lngId = ExecuteInsert( _
            "INSERT INTO [Mouvement] ([TypeMouvement], [VoitureID], [TiersID], [IntermediaireID], [LieuID], " & _
            "[DateMouvement], [Prix], [Commission], [ModeReglement], [Statut], [TypeRetenue], [ValeurRetenue], " & _
            "[DateLimiteSolde], [DateAnnulation], [MotifAnnulation], [UtilisateurID]) " & _
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", _
            pMovement.MovementType, pMovement.CarId, pMovement.PartyId, NullIfZero(pMovement.IntermediaryId), _
            NullIfZero(pMovement.PlaceId), pMovement.MovementDate, pMovement.Price, pMovement.Commission, _
            pMovement.SettlementMode, pMovement.Status, NullIfEmpty(pMovement.RetentionType), _
            pMovement.RetentionValue, pMovement.BalanceDueDate, pMovement.CancelDate, _
            NullIfEmpty(pMovement.CancelReason), pMovement.UserId)
        pMovement.SetId lngId
    Else
        lngRows = ExecuteNonQuery( _
            "UPDATE [Mouvement] SET [TypeMouvement] = ?, [VoitureID] = ?, [TiersID] = ?, [IntermediaireID] = ?, " & _
            "[LieuID] = ?, [DateMouvement] = ?, [Prix] = ?, [Commission] = ?, [ModeReglement] = ?, [Statut] = ?, " & _
            "[TypeRetenue] = ?, [ValeurRetenue] = ?, [DateLimiteSolde] = ?, [DateAnnulation] = ?, " & _
            "[MotifAnnulation] = ?, [UtilisateurID] = ? WHERE [MouvementID] = ?", _
            pMovement.MovementType, pMovement.CarId, pMovement.PartyId, NullIfZero(pMovement.IntermediaryId), _
            NullIfZero(pMovement.PlaceId), pMovement.MovementDate, pMovement.Price, pMovement.Commission, _
            pMovement.SettlementMode, pMovement.Status, NullIfEmpty(pMovement.RetentionType), _
            pMovement.RetentionValue, pMovement.BalanceDueDate, pMovement.CancelDate, _
            NullIfEmpty(pMovement.CancelReason), pMovement.UserId, pMovement.Id)
        If lngRows = 0 Then
            Err.Raise ERR_DATA_NOT_FOUND, "DataAccess.SaveMovement", _
                "Mouvement introuvable (identifiant " & pMovement.Id & ")."
        End If
    End If
End Sub

' ---------- Paiement ----------

' Insère un paiement (un paiement enregistré ne se modifie plus : il s'annule par un remboursement)
Public Sub SavePayment(ByVal pPayment As Payment)
    pPayment.Validate
    If pPayment.Id <> 0 Then
        Err.Raise ERR_PAYMENT_INVALID, "DataAccess.SavePayment", "Un paiement enregistré ne se modifie pas."
    End If
    pPayment.SetId ExecuteInsert( _
        "INSERT INTO [Paiement] ([MouvementID], [ReservationID], [EcheanceID], [FactureID], [NumeroRecu], [Sens], " & _
        "[DatePaiement], [Montant], [ModePaiement], [UtilisateurID]) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", _
        NullIfZero(pPayment.MovementId), NullIfZero(pPayment.ReservationId), Null, NullIfZero(pPayment.InvoiceId), _
        pPayment.ReceiptNumber, pPayment.Direction, pPayment.PaymentDate, pPayment.Amount, pPayment.PaymentMode, _
        pPayment.UserId)
End Sub

' Total des paiements d'un mouvement dans un sens (ENCAISSEMENT ou REMBOURSEMENT)
Public Function PaymentsTotal(ByVal pMovementId As Long, ByVal pDirection As String) As Currency
    Dim vTotal As Variant

    vTotal = QueryValue("SELECT SUM([Montant]) FROM [Paiement] WHERE [MouvementID] = ? AND [Sens] = ?", _
                        pMovementId, pDirection)
    If Not IsNull(vTotal) Then PaymentsTotal = CCur(vTotal)
End Function

' Rattache à la facture les encaissements du mouvement qui n'ont pas encore de facture
Public Sub LinkPaymentsToInvoice(ByVal pMovementId As Long, ByVal pInvoiceId As Long)
    ExecuteNonQuery "UPDATE [Paiement] SET [FactureID] = ? WHERE [MouvementID] = ? AND [Sens] = ? AND [FactureID] IS NULL", _
                    pInvoiceId, pMovementId, PAYMENT_DIRECTION_IN
End Sub

' ---------- Facture ----------

' Facture d'une vente : la facture émise, sinon la dernière (Nothing si aucune)
Public Function LoadInvoiceOfMovement(ByVal pMovementId As Long) As Invoice
    Const SQL_INVOICE As String = _
        "SELECT TOP 1 [FactureID], [NumeroFacture], [MouvementID], [TiersID], [DateFacture], [Designation], " & _
        "[MontantHT], [TauxTaxe], [MontantTaxe], [MontantTTC], [Statut], [FactureOrigineID], [UtilisateurID] " & _
        "FROM [Facture] WHERE [MouvementID] = ? "
    Dim inv As Invoice

    Set inv = InvoiceFromQuery(SQL_INVOICE & "AND [Statut] = ? ORDER BY [FactureID] DESC", pMovementId, INVOICE_STATUS_ISSUED)
    If inv Is Nothing Then Set inv = InvoiceFromQuery(SQL_INVOICE & "ORDER BY [FactureID] DESC", pMovementId)
    Set LoadInvoiceOfMovement = inv
End Function

' Insère la facture ; une facture existante ne peut changer que de statut (annulation)
Public Sub SaveInvoice(ByVal pInvoice As Invoice)
    pInvoice.Validate
    If pInvoice.Id = 0 Then
        pInvoice.SetId ExecuteInsert( _
            "INSERT INTO [Facture] ([NumeroFacture], [MouvementID], [TiersID], [DateFacture], [Designation], [MontantHT], " & _
            "[TauxTaxe], [MontantTaxe], [MontantTTC], [Statut], [FactureOrigineID], [UtilisateurID]) " & _
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", _
            pInvoice.InvoiceNumber, pInvoice.MovementId, pInvoice.PartyId, pInvoice.InvoiceDate, pInvoice.Designation, _
            pInvoice.NetAmount, pInvoice.TaxRate, pInvoice.TaxAmount, pInvoice.TotalAmount, pInvoice.Status, _
            NullIfZero(pInvoice.OriginInvoiceId), pInvoice.UserId)
    ElseIf ExecuteNonQuery("UPDATE [Facture] SET [Statut] = ? WHERE [FactureID] = ?", pInvoice.Status, pInvoice.Id) = 0 Then
        Err.Raise ERR_DATA_NOT_FOUND, "DataAccess.SaveInvoice", "Facture introuvable (identifiant " & pInvoice.Id & ")."
    End If
End Sub

' ---------- Suivi du certificat de vente ----------

' Suivi du CV d'une vente (Nothing si aucun)
Public Function LoadCvTrackingOfMovement(ByVal pMovementId As Long) As CvTracking
    Dim AdoRs As ADODB.Recordset
    Dim cv As CvTracking

    Set AdoRs = QueryRecordset( _
        "SELECT [SuiviCVID], [MouvementID], [NumeroCV], [CommissariatID], [DateDepot], [DateSignature], " & _
        "[FraisPayes], [Statut] FROM [SuiviCV] WHERE [MouvementID] = ?", pMovementId)
    If Not AdoRs.EOF Then
        Set cv = New CvTracking
        cv.Hydrate CLng(AdoRs.Fields("SuiviCVID").Value), CLng(AdoRs.Fields("MouvementID").Value), _
                   AdoRs.Fields("NumeroCV").Value, AdoRs.Fields("CommissariatID").Value, _
                   AdoRs.Fields("DateDepot").Value, AdoRs.Fields("DateSignature").Value, _
                   CCur(AdoRs.Fields("FraisPayes").Value), CStr(AdoRs.Fields("Statut").Value)
    End If
    AdoRs.Close
    Set LoadCvTrackingOfMovement = cv
End Function

' Insère ou met à jour le suivi du CV
Public Sub SaveCvTracking(ByVal pCv As CvTracking)
    If pCv.MovementId <= 0 Then
        Err.Raise ERR_CV, "DataAccess.SaveCvTracking", "Le suivi du CV doit être rattaché à une vente."
    End If
    If pCv.Id = 0 Then
        pCv.SetId ExecuteInsert( _
            "INSERT INTO [SuiviCV] ([MouvementID], [NumeroCV], [CommissariatID], [DateDepot], [DateSignature], " & _
            "[FraisPayes], [Statut]) VALUES (?, ?, ?, ?, ?, ?, ?)", _
            pCv.MovementId, NullIfEmpty(pCv.CvNumber), NullIfZero(pCv.PoliceStationId), pCv.FilingDate, _
            pCv.SignatureDate, pCv.FeesPaid, pCv.Status)
    ElseIf ExecuteNonQuery( _
            "UPDATE [SuiviCV] SET [NumeroCV] = ?, [CommissariatID] = ?, [DateDepot] = ?, [DateSignature] = ?, " & _
            "[FraisPayes] = ?, [Statut] = ? WHERE [SuiviCVID] = ?", _
            NullIfEmpty(pCv.CvNumber), NullIfZero(pCv.PoliceStationId), pCv.FilingDate, pCv.SignatureDate, _
            pCv.FeesPaid, pCv.Status, pCv.Id) = 0 Then
        Err.Raise ERR_DATA_NOT_FOUND, "DataAccess.SaveCvTracking", "Suivi du CV introuvable (identifiant " & pCv.Id & ")."
    End If
End Sub

' ---------- Méthodes privées ----------

' Texte d'une colonne facultative (Null -> chaîne vide)
Private Function TextOf(ByVal pValue As Variant) As String
    If IsNull(pValue) Or IsEmpty(pValue) Then
        TextOf = ""
    Else
        TextOf = CStr(pValue)
    End If
End Function

' Deux voitures ne peuvent pas avoir le même numéro de châssis
Private Sub RequireUniqueChassis(ByVal pCar As Car)
    Dim lngCount As Long

    lngCount = CLng(QueryValue("SELECT COUNT(*) FROM [Voiture] WHERE [Chassis] = ? AND [VoitureID] <> ?", _
                            pCar.Chassis, pCar.Id))
    If lngCount > 0 Then
        Err.Raise ERR_DATA_DUPLICATE, "DataAccess.SaveCar", _
            "Une voiture avec ce numéro de châssis existe déjà : " & pCar.Chassis
    End If
End Sub

' Premier mouvement renvoyé par la requête (Nothing si aucun)
Private Function MovementFromQuery(ByVal pSql As String, ParamArray pParams() As Variant) As Movement
    Dim AdoRs As ADODB.Recordset
    Dim m As Movement

    Set AdoRs = RunSql(pSql, pParams, kxRecordset)
    If Not AdoRs.EOF Then
        Set m = New Movement
        m.Hydrate CLng(AdoRs.Fields("MouvementID").Value), CStr(AdoRs.Fields("TypeMouvement").Value), _
                  CLng(AdoRs.Fields("VoitureID").Value), CLng(AdoRs.Fields("TiersID").Value), _
                  AdoRs.Fields("IntermediaireID").Value, AdoRs.Fields("LieuID").Value, _
                  CDate(AdoRs.Fields("DateMouvement").Value), CCur(AdoRs.Fields("Prix").Value), _
                  CCur(AdoRs.Fields("Commission").Value), CStr(AdoRs.Fields("ModeReglement").Value), _
                  CStr(AdoRs.Fields("Statut").Value), AdoRs.Fields("TypeRetenue").Value, _
                  AdoRs.Fields("ValeurRetenue").Value, AdoRs.Fields("DateLimiteSolde").Value, _
                  AdoRs.Fields("DateAnnulation").Value, AdoRs.Fields("MotifAnnulation").Value, _
                  CLng(AdoRs.Fields("UtilisateurID").Value)
    End If
    AdoRs.Close
    Set MovementFromQuery = m
End Function

' Identifiant facultatif : 0 -> Null
Private Function NullIfZero(ByVal pId As Long) As Variant
    If pId = 0 Then NullIfZero = Null Else NullIfZero = pId
End Function

' Première facture renvoyée par la requête (Nothing si aucune)
Private Function InvoiceFromQuery(ByVal pSql As String, ParamArray pParams() As Variant) As Invoice
    Dim AdoRs As ADODB.Recordset
    Dim inv As Invoice

    Set AdoRs = RunSql(pSql, pParams, kxRecordset)
    If Not AdoRs.EOF Then
        Set inv = New Invoice
        inv.Hydrate CLng(AdoRs.Fields("FactureID").Value), CStr(AdoRs.Fields("NumeroFacture").Value), _
                    CLng(AdoRs.Fields("MouvementID").Value), CLng(AdoRs.Fields("TiersID").Value), _
                    CDate(AdoRs.Fields("DateFacture").Value), CStr(AdoRs.Fields("Designation").Value), _
                    CCur(AdoRs.Fields("MontantHT").Value), CDbl(AdoRs.Fields("TauxTaxe").Value), _
                    CCur(AdoRs.Fields("MontantTaxe").Value), CCur(AdoRs.Fields("MontantTTC").Value), _
                    CStr(AdoRs.Fields("Statut").Value), AdoRs.Fields("FactureOrigineID").Value, _
                    CLng(AdoRs.Fields("UtilisateurID").Value)
    End If
    AdoRs.Close
    Set InvoiceFromQuery = inv
End Function
