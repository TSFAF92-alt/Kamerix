VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Car"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Classe Car (voiture) : porte le statut (transitions contrôlées) et les données descriptives (table Voiture).
' Les permissions (annuler une vente, retirer) sont vérifiées par la couche Session, pas ici.
' Une voiture neuve (New Car) démarre au statut EN_STOCK.

Private Const MAX_COLOR_LENGTH As Long = 30
Private Const MAX_CHASSIS_LENGTH As Long = 30
Private Const MAX_PLATE_LENGTH As Long = 20
Private Const MIN_MODEL_YEAR As Long = 1950
Private Const MAX_REASON_LENGTH As Long = 100

Private mIdentifier As Long
Private mStatus As String
Private mWithdrawalReason As String
Private mModelId As Long
Private mModelYear As Variant          ' Null = non renseignée
Private mColor As String
Private mChassis As String
Private mPlateNumber As String
Private mMileage As Variant    ' Null = non renseigné

Private Sub Class_Initialize()
    mStatus = CAR_STATUS_IN_STOCK
    mModelYear = Null
    mMileage = Null
End Sub

' ---------- Identité et statut ----------

Public Property Get Id() As Long
    Id = mIdentifier
End Property

Public Property Get Status() As String
    Status = mStatus
End Property

Public Property Get WithdrawalReason() As String
    WithdrawalReason = mWithdrawalReason
End Property

' ---------- Données descriptives ----------

Public Property Get ModelId() As Long
    ModelId = mModelId
End Property
Public Property Let ModelId(ByVal pValue As Long)
    RequireEditable "ModelId"
    mModelId = pValue
End Property

Public Property Get ModelYear() As Variant
    ModelYear = mModelYear
End Property
Public Property Let ModelYear(ByVal pValue As Variant)
    RequireEditable "ModelYear"
    mModelYear = NumberOrNull(pValue)
End Property

Public Property Get Color() As String
    Color = mColor
End Property
Public Property Let Color(ByVal pValue As String)
    RequireEditable "Color"
    mColor = Trim$(pValue)
End Property

Public Property Get Chassis() As String
    Chassis = mChassis
End Property
Public Property Let Chassis(ByVal pValue As String)
    RequireEditable "Chassis"
    mChassis = UCase$(Trim$(pValue))
End Property

Public Property Get PlateNumber() As String
    PlateNumber = mPlateNumber
End Property
Public Property Let PlateNumber(ByVal pValue As String)
    RequireEditable "PlateNumber"
    mPlateNumber = UCase$(Trim$(pValue))
End Property

Public Property Get Mileage() As Variant
    Mileage = mMileage
End Property
Public Property Let Mileage(ByVal pValue As Variant)
    RequireEditable "Mileage"
    mMileage = NumberOrNull(pValue)
End Property

' ---------- Chargement depuis la base (classe d'accès aux données uniquement) ----------

' Charge le statut sans contrôle de transition
Friend Sub Hydrate(ByVal pId As Long, ByVal pStatus As String, Optional ByVal pWithdrawalReason As String = "")
    If Not IsValidStatus(pStatus) Then
        Err.Raise ERR_CAR_STATUS, "Car.Hydrate", "Statut inconnu : " & pStatus
    End If
    mIdentifier = pId
    mStatus = pStatus
    mWithdrawalReason = pWithdrawalReason
End Sub

' Charge les données descriptives lues dans la base (colonnes vides en Null)
Friend Sub HydrateData(ByVal pModelId As Long, ByVal pModelYear As Variant, ByVal pColor As Variant, _
                           ByVal pChassis As Variant, ByVal pPlateNumber As Variant, _
                           ByVal pMileage As Variant)
    mModelId = pModelId
    mModelYear = NumberOrNull(pModelYear)
    mColor = TextOf(pColor)
    mChassis = UCase$(TextOf(pChassis))
    mPlateNumber = UCase$(TextOf(pPlateNumber))
    mMileage = NumberOrNull(pMileage)
End Sub

' Appelé par la classe d'accès aux données après l'insertion
Friend Sub SetId(ByVal pId As Long)
    mIdentifier = pId
End Sub

' ---------- Validation des données ----------

' Liste de tous les problèmes trouvés (collection vide = données valides)
Public Function ValidationErrors() As Collection
    Dim e As Collection
    Set e = New Collection

    If mModelId <= 0 Then
        e.Add "Le modèle est obligatoire."
    End If

    If Len(mChassis) = 0 Then
        e.Add "Le numéro de châssis est obligatoire."
    ElseIf Len(mChassis) > MAX_CHASSIS_LENGTH Then
        e.Add "Le châssis ne doit pas dépasser " & MAX_CHASSIS_LENGTH & " caractères."
    ElseIf Not IsValidChassis(mChassis) Then
        e.Add "Le châssis ne peut contenir que des lettres, des chiffres et des tirets."
    End If

    If Not IsNull(mModelYear) Then
        If Not IsNumeric(mModelYear) Then
            e.Add "L'année doit être un nombre."
        ElseIf CDbl(mModelYear) < MIN_MODEL_YEAR Or CDbl(mModelYear) > Year(Date) + 1 Then
            e.Add "L'année doit être comprise entre " & MIN_MODEL_YEAR & " et " & (Year(Date) + 1) & "."
        End If
    End If

    If Not IsNull(mMileage) Then
        If Not IsNumeric(mMileage) Then
            e.Add "Le kilométrage doit être un nombre."
        ElseIf CDbl(mMileage) < 0 Then
            e.Add "Le kilométrage ne peut pas être négatif."
        End If
    End If

    If Len(mColor) > MAX_COLOR_LENGTH Then
        e.Add "La couleur ne doit pas dépasser " & MAX_COLOR_LENGTH & " caractères."
    End If
    If Len(mPlateNumber) > MAX_PLATE_LENGTH Then
        e.Add "L'immatriculation ne doit pas dépasser " & MAX_PLATE_LENGTH & " caractères."
    End If

    Set ValidationErrors = e
End Function

' Lève une erreur unique regroupant tous les problèmes (à appeler avant l'enregistrement)
Public Sub Validate()
    Dim e As Collection
    Dim msg As String
    Dim i As Long

    Set e = ValidationErrors()
    If e.Count = 0 Then Exit Sub

    For i = 1 To e.Count
        If i > 1 Then msg = msg & vbCrLf
        msg = msg & "- " & e.Item(i)
    Next i
    Err.Raise ERR_CAR_DATA, "Car.Validate", msg
End Sub

' ---------- Transitions de statut ----------

' Indique si la transition vers le statut cible est autorisée depuis le statut actuel
Public Function CanTransitionTo(ByVal pTarget As String) As Boolean
    Select Case mStatus
        Case CAR_STATUS_IN_STOCK
            CanTransitionTo = (pTarget = CAR_STATUS_RESERVED Or _
                           pTarget = CAR_STATUS_SALE_IN_PROGRESS Or _
                           pTarget = CAR_STATUS_WITHDRAWN)
        Case CAR_STATUS_RESERVED
            CanTransitionTo = (pTarget = CAR_STATUS_IN_STOCK Or _
                           pTarget = CAR_STATUS_SALE_IN_PROGRESS)
        Case CAR_STATUS_SALE_IN_PROGRESS
            CanTransitionTo = (pTarget = CAR_STATUS_SOLD Or _
                           pTarget = CAR_STATUS_IN_STOCK)
        Case Else
            CanTransitionTo = False   ' VENDU et RETIRE sont définitifs
    End Select
End Function

' En stock -> Réservé : un client verse une avance
Public Sub Reserve()
    TransitionTo CAR_STATUS_RESERVED, "Reserve"
End Sub

' Réservé -> En stock : désistement du client ou expiration de la date limite
Public Sub CancelReservation()
    RequireStatus CAR_STATUS_RESERVED, "CancelReservation"
    TransitionTo CAR_STATUS_IN_STOCK, "CancelReservation"
End Sub

' Réservé -> Vente en cours : le client a payé le solde avant la date limite
Public Sub ConfirmSale()
    RequireStatus CAR_STATUS_RESERVED, "ConfirmSale"
    TransitionTo CAR_STATUS_SALE_IN_PROGRESS, "ConfirmSale"
End Sub

' En stock -> Vente en cours : vente directe, sans réservation (comptant ou échelonnée)
Public Sub StartDirectSale()
    RequireStatus CAR_STATUS_IN_STOCK, "StartDirectSale"
    TransitionTo CAR_STATUS_SALE_IN_PROGRESS, "StartDirectSale"
End Sub

' Vente en cours -> Vendu : le CV est signé (état définitif)
Public Sub MarkAsSold()
    TransitionTo CAR_STATUS_SOLD, "MarkAsSold"
End Sub

' Vente en cours -> En stock : vente annulée avant la signature du CV
' (motif obligatoire ; il est conservé par le mouvement, dans Mouvement.MotifAnnulation, pas par la voiture)
Public Sub CancelSale(ByVal pReason As String)
    RequireReason pReason, "CancelSale"
    TransitionTo CAR_STATUS_IN_STOCK, "CancelSale"
End Sub

' En stock -> Retiré : sortie du stock (motif obligatoire, conservé dans la colonne Voiture.MotifRetrait)
Public Sub Withdraw(ByVal pReason As String)
    RequireReason pReason, "Withdraw"
    TransitionTo CAR_STATUS_WITHDRAWN, "Withdraw"
    mWithdrawalReason = Trim$(pReason)
End Sub

' ---------- Méthodes privées ----------

Private Sub TransitionTo(ByVal pTarget As String, ByVal pSource As String)
    If Not CanTransitionTo(pTarget) Then
        Err.Raise ERR_CAR_TRANSITION, "Car." & pSource, _
            "Transition interdite : " & mStatus & " -> " & pTarget
    End If
    mStatus = pTarget
End Sub

Private Sub RequireStatus(ByVal pExpected As String, ByVal pSource As String)
    If mStatus <> pExpected Then
        Err.Raise ERR_CAR_TRANSITION, "Car." & pSource, _
            "Cette opération exige le statut " & pExpected & " (statut actuel : " & mStatus & ")"
    End If
End Sub

Private Sub RequireReason(ByVal pReason As String, ByVal pSource As String)
    If Len(Trim$(pReason)) = 0 Then
        Err.Raise ERR_CAR_REASON, "Car." & pSource, "Un motif est obligatoire."
    ElseIf Len(Trim$(pReason)) > MAX_REASON_LENGTH Then
        Err.Raise ERR_CAR_REASON, "Car." & pSource, _
            "Le motif ne doit pas dépasser " & MAX_REASON_LENGTH & " caractères."
    End If
End Sub

' Une voiture vendue ou retirée est figée : ses données ne changent plus
Private Sub RequireEditable(ByVal pSource As String)
    If mStatus = CAR_STATUS_SOLD Or mStatus = CAR_STATUS_WITHDRAWN Then
        Err.Raise ERR_CAR_LOCKED, "Car." & pSource, _
            "Une voiture vendue ou retirée ne peut plus être modifiée."
    End If
End Sub

Private Function IsValidStatus(ByVal pStatus As String) As Boolean
    Select Case pStatus
        Case CAR_STATUS_IN_STOCK, CAR_STATUS_RESERVED, CAR_STATUS_SALE_IN_PROGRESS, _
             CAR_STATUS_SOLD, CAR_STATUS_WITHDRAWN
            IsValidStatus = True
    End Select
End Function

Private Function IsValidChassis(ByVal pText As String) As Boolean
    Dim i As Long
    Dim c As String

    For i = 1 To Len(pText)
        c = Mid$(pText, i, 1)
        If Not ((c >= "A" And c <= "Z") Or (c >= "0" And c <= "9") Or c = "-") Then Exit Function
    Next i
    IsValidChassis = True
End Function

' Vide ou Null -> Null ; nombre entier -> Long ; autre valeur conservée telle quelle (la validation la signalera)
Private Function NumberOrNull(ByVal pValue As Variant) As Variant
    If IsNull(pValue) Or IsEmpty(pValue) Then
        NumberOrNull = Null
    ElseIf VarType(pValue) = vbString And Len(Trim$(CStr(pValue))) = 0 Then
        NumberOrNull = Null
    ElseIf IsNumeric(pValue) Then
        If Abs(CDbl(pValue)) < 2000000000# Then
            NumberOrNull = CLng(pValue)
        Else
            NumberOrNull = pValue
        End If
    Else
        NumberOrNull = pValue
    End If
End Function

Private Function TextOf(ByVal pValue As Variant) As String
    If IsNull(pValue) Or IsEmpty(pValue) Then
        TextOf = ""
    Else
        TextOf = Trim$(CStr(pValue))
    End If
End Function
