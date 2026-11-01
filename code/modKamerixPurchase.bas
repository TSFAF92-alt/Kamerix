Attribute VB_Name = "modKamerixPurchase"
Option Explicit
Option Private Module

'==============================================================================
' KAMERIX - Service d'achat : enregistrer l'achat d'une voiture
'------------------------------------------------------------------------------
' Role        : enregistrer, en une seule transaction (tout ou rien), la voiture
'               (nouvelle ou deja saisie) et son mouvement d'achat, solde des
'               sa creation (un achat est regle comptant au vendeur).
' Droits      : ecriture sur le module Voitures (Administrateur, Responsable).
' Regles      : - le vendeur existe et est marque "vendeur" (Party.IsSeller) ;
'               - la voiture est en stock et n'a pas deja un achat ;
'               - prix > 0 ; commission >= 0, inferieure au prix, et seulement
'                 avec un intermediaire ; l'intermediaire et le lieu existent ;
'               - l'utilisateur connecte est inscrit sur le mouvement.
' Transaction : ouverte par le service, sauf si l'appelant en a deja une (ses
'               ecritures en font alors partie). Toutes les donnees sont
'               validees avant la premiere ecriture.
' Cout        : CarCostPrice = prix d'achat + commission d'achat + charges
'               rattachees a la voiture (decide le 31/10/2026).
' Erreurs     : ERR_PURCHASE (regles), ERR_MOVEMENT_INVALID (donnees), ERR_SESSION_*.
' Dependances : DataAccess, Car, Party, Movement, modKamerixSession, modKamerixSql.
'==============================================================================

' Enregistre l'achat d'une voiture et retourne l'identifiant du mouvement.
'   pCar          : voiture a enregistrer (Id = 0) ou deja enregistree, au statut EN_STOCK
'   pSellerId     : tiers vendeur
'   pPrice        : prix d'achat
'   pCommission   : commission de l'intermediaire (0 si aucun)
'   pIntermediaryId, pPlaceId : 0 si aucun
'   pDate         : date de l'achat (aujourd'hui par defaut)
Public Function RecordPurchase(ByVal pCar As Car, ByVal pSellerId As Long, ByVal pPrice As Currency, _
                               Optional ByVal pCommission As Currency = 0, Optional ByVal pIntermediaryId As Long = 0, _
                               Optional ByVal pPlaceId As Long = 0, Optional ByVal pDate As Variant) As Long
    Const SRC As String = "modKamerixPurchase.RecordPurchase"
    Dim objData As DataAccess
    Dim objSeller As Party
    Dim objPurchase As Movement
    Dim blnNewCar As Boolean
    Dim blnOwnTransaction As Boolean
    Dim lngErr As Long
    Dim strErr As String
    Dim strSrc As String

    RequireWrite MOD_CARS
    Set objData = New DataAccess

    ' ---------- Regles controlees avant toute ecriture ----------
    If pCar Is Nothing Then Err.Raise ERR_PURCHASE, SRC, "La voiture est obligatoire."
    If pCar.Status <> CAR_STATUS_IN_STOCK Then
        Err.Raise ERR_PURCHASE, SRC, "Seule une voiture en stock peut " & ChrW(234) & "tre achet" & ChrW(233) & "e (statut : " & pCar.Status & ")."
    End If
    If pCar.Id <> 0 Then
        If Not objData.LoadPurchaseOfCar(pCar.Id) Is Nothing Then
            Err.Raise ERR_PURCHASE, SRC, "Cette voiture a d" & ChrW(233) & "j" & ChrW(224) & " un achat enregistr" & ChrW(233) & "."
        End If
    End If

    Set objSeller = objData.LoadParty(pSellerId)
    If objSeller Is Nothing Then
        Err.Raise ERR_PURCHASE, SRC, "Vendeur introuvable (identifiant " & pSellerId & ")."
    End If
    If Not objSeller.IsSeller Then
        Err.Raise ERR_PURCHASE, SRC, objSeller.Name & " n'est pas enregistr" & ChrW(233) & " comme vendeur."
    End If
    If pIntermediaryId <> 0 Then
        If objData.LoadIntermediary(pIntermediaryId) Is Nothing Then
            Err.Raise ERR_PURCHASE, SRC, "Interm" & ChrW(233) & "diaire introuvable (identifiant " & pIntermediaryId & ")."
        End If
    End If
    If pPlaceId <> 0 Then
        If Len(PlaceTypeOf(pPlaceId)) = 0 Then
            Err.Raise ERR_PURCHASE, SRC, "Lieu introuvable (identifiant " & pPlaceId & ")."
        End If
    End If

    Set objPurchase = New Movement
    objPurchase.MovementType = MOVEMENT_TYPE_PURCHASE
    objPurchase.SettlementMode = SETTLEMENT_CASH
    objPurchase.PartyId = pSellerId
    objPurchase.Price = pPrice
    objPurchase.Commission = pCommission
    objPurchase.IntermediaryId = pIntermediaryId
    objPurchase.PlaceId = pPlaceId
    If Not IsMissing(pDate) Then objPurchase.MovementDate = CDate(pDate)
    objPurchase.SetUserId CurrentUserId()
    objPurchase.MarkSettled

    ' Tout est valide AVANT la premiere ecriture : si l'appelant a deja ouvert une
    ' transaction, ce service ne l'annule pas lui-meme, rien ne doit y rester.
    pCar.Validate
    objPurchase.CarId = pCar.Id
    objPurchase.Validate pCarPending:=(pCar.Id = 0)

    ' ---------- Ecritures : tout ou rien ----------
    blnNewCar = (pCar.Id = 0)
    blnOwnTransaction = Not TransactionActive()
    If blnOwnTransaction Then BeginTransaction
    On Error GoTo EH

    objData.SaveCar pCar
    objPurchase.CarId = pCar.Id
    objData.SaveMovement objPurchase

    If blnOwnTransaction Then CommitTransaction
    RecordPurchase = objPurchase.Id
    Exit Function

EH:
    lngErr = Err.Number
    strErr = Err.Description
    strSrc = Err.Source
    If blnOwnTransaction Then RollbackTransaction
    If blnNewCar Then pCar.SetId 0          ' la voiture n'a pas ete enregistree
    Err.Raise lngErr, strSrc, strErr
End Function

' Achat d'une voiture (Nothing si aucun). Lecture : module Voitures.
Public Function PurchaseOfCar(ByVal pCarId As Long) As Movement
    Dim objData As DataAccess

    RequireRead MOD_CARS
    Set objData = New DataAccess
    Set PurchaseOfCar = objData.LoadPurchaseOfCar(pCarId)
End Function

' Cout de revient : prix d'achat + commission versee a l'achat + charges rattachees a la voiture.
' Erreur ERR_PURCHASE si la voiture n'a pas d'achat. Lecture : module Voitures.
Public Function CarCostPrice(ByVal pCarId As Long) As Currency
    Dim objPurchase As Movement
    Dim vCharges As Variant

    Set objPurchase = PurchaseOfCar(pCarId)
    If objPurchase Is Nothing Then
        Err.Raise ERR_PURCHASE, "modKamerixPurchase.CarCostPrice", _
            "Cette voiture n'a pas d'achat enregistr" & ChrW(233) & " : co" & ChrW(251) & "t de revient inconnu."
    End If

    vCharges = QueryValue("SELECT SUM([Montant]) FROM [Charge] WHERE [VoitureID] = ?", pCarId)
    If IsNull(vCharges) Then vCharges = 0
    CarCostPrice = objPurchase.Price + objPurchase.Commission + CCur(vCharges)
End Function
