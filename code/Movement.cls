VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Movement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Classe Movement : un achat ou une vente d'une voiture (table Mouvement).
' Statut : EN_COURS à la création, puis SOLDE (prix entièrement payé) ou ANNULE.
' Un achat est enregistré directement soldé. Les règles qui demandent la base
' (vendeur réel, un seul achat par voiture, droits) sont contrôlées par le
' service qui crée le mouvement, pas ici.
' Identifiants facultatifs : 0 = aucun (intermédiaire, lieu).

Private Const MAX_REASON_LENGTH As Long = 100

Private mIdentifier As Long
Private mMovementType As String
Private mCarId As Long
Private mPartyId As Long
Private mIntermediaryId As Long
Private mPlaceId As Long
Private mMovementDate As Date
Private mPrice As Currency
Private mCommission As Currency
Private mSettlementMode As String
Private mStatus As String
Private mRetentionType As String
Private mRetentionValue As Variant      ' Null = non renseignée
Private mBalanceDueDate As Variant      ' Null = non renseignée
Private mCancelDate As Variant          ' Null tant que le mouvement n'est pas annulé
Private mCancelReason As String
Private mUserId As Long

Private Sub Class_Initialize()
    mStatus = MOVEMENT_STATUS_IN_PROGRESS
    mSettlementMode = SETTLEMENT_CASH
    mMovementDate = Date
    mRetentionValue = Null
    mBalanceDueDate = Null
    mCancelDate = Null
End Sub

' ---------- Identité et statut ----------

Public Property Get Id() As Long
    Id = mIdentifier
End Property

Public Property Get Status() As String
    Status = mStatus
End Property

Public Property Get CancelDate() As Variant
    CancelDate = mCancelDate
End Property

Public Property Get CancelReason() As String
    CancelReason = mCancelReason
End Property

' Utilisateur qui a enregistré le mouvement (fixé par le service, depuis la session)
Public Property Get UserId() As Long
    UserId = mUserId
End Property

' ---------- Données ----------

' ACHAT ou VENTE
Public Property Get MovementType() As String
    MovementType = mMovementType
End Property
Public Property Let MovementType(ByVal pValue As String)
    mMovementType = UCase$(Trim$(pValue))
End Property

Public Property Get CarId() As Long
    CarId = mCarId
End Property
Public Property Let CarId(ByVal pValue As Long)
    mCarId = pValue
End Property

' Vendeur (achat) ou acheteur (vente)
Public Property Get PartyId() As Long
    PartyId = mPartyId
End Property
Public Property Let PartyId(ByVal pValue As Long)
    mPartyId = pValue
End Property

Public Property Get IntermediaryId() As Long
    IntermediaryId = mIntermediaryId
End Property
Public Property Let IntermediaryId(ByVal pValue As Long)
    mIntermediaryId = pValue
End Property

Public Property Get PlaceId() As Long
    PlaceId = mPlaceId
End Property
Public Property Let PlaceId(ByVal pValue As Long)
    mPlaceId = pValue
End Property

Public Property Get MovementDate() As Date
    MovementDate = mMovementDate
End Property
Public Property Let MovementDate(ByVal pValue As Date)
    mMovementDate = pValue
End Property

' Prix convenu, TTC
Public Property Get Price() As Currency
    Price = mPrice
End Property
Public Property Let Price(ByVal pValue As Currency)
    mPrice = pValue
End Property

' Commission de l'intermédiaire (coût interne, jamais facturée au client)
Public Property Get Commission() As Currency
    Commission = mCommission
End Property
Public Property Let Commission(ByVal pValue As Currency)
    mCommission = pValue
End Property

' COMPTANT ou ECHELONNE
Public Property Get SettlementMode() As String
    SettlementMode = mSettlementMode
End Property
Public Property Let SettlementMode(ByVal pValue As String)
    mSettlementMode = UCase$(Trim$(pValue))
End Property

' Retenue et date limite du solde (vente échelonnée, phase 6)
Public Property Get RetentionType() As String
    RetentionType = mRetentionType
End Property
Public Property Get RetentionValue() As Variant
    RetentionValue = mRetentionValue
End Property
Public Property Get BalanceDueDate() As Variant
    BalanceDueDate = mBalanceDueDate
End Property

' ---------- Changements de statut ----------

' EN_COURS -> SOLDE : le prix est entièrement payé (un achat est soldé dès sa création)
Public Sub MarkSettled()
    If mStatus <> MOVEMENT_STATUS_IN_PROGRESS Then
        Err.Raise ERR_MOVEMENT_STATUS, "Movement.MarkSettled", _
            "Seul un mouvement en cours peut être soldé (statut actuel : " & mStatus & ")."
    End If
    mStatus = MOVEMENT_STATUS_SETTLED
End Sub

' EN_COURS ou SOLDE -> ANNULE, avec date et motif (le service vérifie le droit et le CV)
Public Sub Cancel(ByVal pReason As String, Optional ByVal pDate As Variant)
    If mStatus = MOVEMENT_STATUS_CANCELLED Then
        Err.Raise ERR_MOVEMENT_STATUS, "Movement.Cancel", "Ce mouvement est déjà annulé."
    End If
    If Len(Trim$(pReason)) = 0 Then
        Err.Raise ERR_MOVEMENT_INVALID, "Movement.Cancel", "Le motif de l'annulation est obligatoire."
    End If
    If Len(Trim$(pReason)) > MAX_REASON_LENGTH Then
        Err.Raise ERR_MOVEMENT_INVALID, "Movement.Cancel", _
            "Le motif ne doit pas dépasser " & MAX_REASON_LENGTH & " caractères."
    End If
    mStatus = MOVEMENT_STATUS_CANCELLED
    mCancelReason = Trim$(pReason)
    If IsMissing(pDate) Then mCancelDate = Now Else mCancelDate = CDate(pDate)
End Sub

' ---------- Chargement depuis la base (classe d'accès aux données uniquement) ----------

Friend Sub Hydrate(ByVal pId As Long, ByVal pMovementType As String, ByVal pCarId As Long, ByVal pPartyId As Long, _
                   ByVal pIntermediaryId As Variant, ByVal pPlaceId As Variant, ByVal pMovementDate As Date, _
                   ByVal pPrice As Currency, ByVal pCommission As Currency, ByVal pSettlementMode As String, _
                   ByVal pStatus As String, ByVal pRetentionType As Variant, ByVal pRetentionValue As Variant, _
                   ByVal pBalanceDueDate As Variant, ByVal pCancelDate As Variant, ByVal pCancelReason As Variant, _
                   ByVal pUserId As Long)
    mIdentifier = pId
    mMovementType = pMovementType
    mCarId = pCarId
    mPartyId = pPartyId
    mIntermediaryId = IdOrZero(pIntermediaryId)
    mPlaceId = IdOrZero(pPlaceId)
    mMovementDate = pMovementDate
    mPrice = pPrice
    mCommission = pCommission
    mSettlementMode = pSettlementMode
    mStatus = pStatus
    mRetentionType = TextOfValue(pRetentionType)
    mRetentionValue = pRetentionValue
    mBalanceDueDate = pBalanceDueDate
    mCancelDate = pCancelDate
    mCancelReason = TextOfValue(pCancelReason)
    mUserId = pUserId
End Sub

' Appelé par la classe d'accès aux données après l'insertion
Friend Sub SetId(ByVal pId As Long)
    mIdentifier = pId
End Sub

' Appelé par le service avant l'enregistrement : utilisateur connecté
Friend Sub SetUserId(ByVal pUserId As Long)
    mUserId = pUserId
End Sub

' ---------- Validation ----------

' Liste de tous les problèmes trouvés (collection vide = mouvement valide).
' pCarPending = Vrai : la voiture sera enregistrée dans la même transaction,
' son identifiant n'est pas encore connu (contrôle fait avant toute écriture).
Public Function ValidationErrors(Optional ByVal pCarPending As Boolean = False) As Collection
    Dim e As Collection

    Set e = New Collection

    If mMovementType <> MOVEMENT_TYPE_PURCHASE And mMovementType <> MOVEMENT_TYPE_SALE Then
        e.Add "Le type de mouvement doit être ACHAT ou VENTE."
    End If
    If mCarId <= 0 And Not pCarPending Then e.Add "La voiture est obligatoire."
    If mPartyId <= 0 Then
        If mMovementType = MOVEMENT_TYPE_PURCHASE Then e.Add "Le vendeur est obligatoire." Else e.Add "L'acheteur est obligatoire."
    End If
    If mPrice <= 0 Then e.Add "Le prix doit être supérieur à zéro."
    If mCommission < 0 Then e.Add "La commission ne peut pas être négative."
    If mCommission > 0 And mIntermediaryId <= 0 Then e.Add "Une commission exige un intermédiaire."
    If mCommission >= mPrice And mPrice > 0 Then e.Add "La commission doit être inférieure au prix."
    If mSettlementMode <> SETTLEMENT_CASH And mSettlementMode <> SETTLEMENT_INSTALLMENT Then
        e.Add "Le mode de règlement doit être COMPTANT ou ECHELONNE."
    End If
    If mMovementType = MOVEMENT_TYPE_PURCHASE And mSettlementMode <> SETTLEMENT_CASH Then
        e.Add "Un achat est toujours réglé comptant."
    End If
    If mMovementDate = 0 Then e.Add "La date du mouvement est obligatoire."
    If mMovementDate > Date + 1 Then e.Add "La date du mouvement ne peut pas être dans le futur."
    If mUserId <= 0 Then e.Add "L'utilisateur qui enregistre le mouvement est inconnu."
    If mStatus = MOVEMENT_STATUS_CANCELLED And (IsNull(mCancelDate) Or Len(mCancelReason) = 0) Then
        e.Add "Un mouvement annulé doit avoir une date et un motif d'annulation."
    End If

    Set ValidationErrors = e
End Function

' Lève une erreur unique regroupant tous les problèmes (à appeler avant l'enregistrement)
Public Sub Validate(Optional ByVal pCarPending As Boolean = False)
    Dim e As Collection
    Dim msg As String
    Dim i As Long

    Set e = ValidationErrors(pCarPending)
    If e.Count = 0 Then Exit Sub

    For i = 1 To e.Count
        If i > 1 Then msg = msg & vbCrLf
        msg = msg & "- " & e.Item(i)
    Next i
    Err.Raise ERR_MOVEMENT_INVALID, "Movement.Validate", msg
End Sub

' ---------- Méthodes privées ----------

Private Function IdOrZero(ByVal pValue As Variant) As Long
    If IsNull(pValue) Or IsEmpty(pValue) Then IdOrZero = 0 Else IdOrZero = CLng(pValue)
End Function
