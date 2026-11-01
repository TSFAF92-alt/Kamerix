VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Payment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Classe Payment : un encaissement ou un remboursement (table Paiement).
' Chaque paiement a son numéro de pièce (REC-… pour un encaissement, RMB-… pour
' un remboursement), attribué par le service dans la transaction qui l'enregistre.
' Un paiement est lié à un mouvement ou à une réservation (au moins un).
' Identifiants facultatifs : 0 = aucun.

Private mIdentifier As Long
Private mMovementId As Long
Private mReservationId As Long
Private mInvoiceId As Long
Private mReceiptNumber As String
Private mDirection As String
Private mPaymentDate As Date
Private mAmount As Currency
Private mPaymentMode As String
Private mUserId As Long

Private Sub Class_Initialize()
    mDirection = PAYMENT_DIRECTION_IN
    mPaymentDate = Date
End Sub

Public Property Get Id() As Long
    Id = mIdentifier
End Property

Public Property Get MovementId() As Long
    MovementId = mMovementId
End Property
Public Property Let MovementId(ByVal pValue As Long)
    mMovementId = pValue
End Property

Public Property Get ReservationId() As Long
    ReservationId = mReservationId
End Property
Public Property Let ReservationId(ByVal pValue As Long)
    mReservationId = pValue
End Property

' Facture à laquelle le paiement est rattaché (renseignée au solde)
Public Property Get InvoiceId() As Long
    InvoiceId = mInvoiceId
End Property

Public Property Get ReceiptNumber() As String
    ReceiptNumber = mReceiptNumber
End Property

' ENCAISSEMENT ou REMBOURSEMENT
Public Property Get Direction() As String
    Direction = mDirection
End Property
Public Property Let Direction(ByVal pValue As String)
    mDirection = UCase$(Trim$(pValue))
End Property

Public Property Get PaymentDate() As Date
    PaymentDate = mPaymentDate
End Property
Public Property Let PaymentDate(ByVal pValue As Date)
    mPaymentDate = pValue
End Property

Public Property Get Amount() As Currency
    Amount = mAmount
End Property
Public Property Let Amount(ByVal pValue As Currency)
    mAmount = pValue
End Property

' Un des modes du paramètre MODES_PAIEMENT
Public Property Get PaymentMode() As String
    PaymentMode = mPaymentMode
End Property
Public Property Let PaymentMode(ByVal pValue As String)
    mPaymentMode = Trim$(pValue)
End Property

Public Property Get UserId() As Long
    UserId = mUserId
End Property

' ---------- Réservé au service et à la classe d'accès aux données ----------

Friend Sub SetReceiptNumber(ByVal pNumber As String)
    mReceiptNumber = pNumber
End Sub

Friend Sub SetUserId(ByVal pUserId As Long)
    mUserId = pUserId
End Sub

Friend Sub SetInvoiceId(ByVal pInvoiceId As Long)
    mInvoiceId = pInvoiceId
End Sub

Friend Sub SetId(ByVal pId As Long)
    mIdentifier = pId
End Sub

' ---------- Validation ----------

Public Function ValidationErrors() As Collection
    Dim e As Collection

    Set e = New Collection
    If mMovementId <= 0 And mReservationId <= 0 Then e.Add "Le paiement doit être lié à un mouvement ou à une réservation."
    If mAmount <= 0 Then e.Add "Le montant doit être supérieur à zéro."
    If mDirection <> PAYMENT_DIRECTION_IN And mDirection <> PAYMENT_DIRECTION_REFUND Then
        e.Add "Le sens du paiement doit être ENCAISSEMENT ou REMBOURSEMENT."
    End If
    If Len(mPaymentMode) = 0 Then
        e.Add "Le mode de paiement est obligatoire."
    ElseIf Not IsValidPaymentMode(mPaymentMode) Then
        e.Add "Mode de paiement inconnu : " & mPaymentMode & "."
    End If
    If mPaymentDate = 0 Then e.Add "La date du paiement est obligatoire."
    If mPaymentDate > Date + 1 Then e.Add "La date du paiement ne peut pas être dans le futur."
    If Len(mReceiptNumber) = 0 Then e.Add "Le numéro de pièce est obligatoire."
    If mUserId <= 0 Then e.Add "L'utilisateur qui enregistre le paiement est inconnu."
    Set ValidationErrors = e
End Function

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
    Err.Raise ERR_PAYMENT_INVALID, "Payment.Validate", msg
End Sub
