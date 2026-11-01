VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Invoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Classe Invoice : facture d'une vente (table Facture).
' Une seule facture par vente, pour la voiture uniquement, émise au solde.
' Montants : le prix convenu est TTC ; HT et taxe en sont déduits
' (modKamerixSettings.SplitTaxIncluded), HT + taxe = TTC.
' Une facture émise ne se modifie plus : elle s'annule (et se remplace).

Private Const MAX_DESIGNATION_LENGTH As Long = 255

Private mIdentifier As Long
Private mInvoiceNumber As String
Private mMovementId As Long
Private mPartyId As Long
Private mInvoiceDate As Date
Private mDesignation As String
Private mNetAmount As Currency
Private mTaxRate As Double
Private mTaxAmount As Currency
Private mTotalAmount As Currency
Private mStatus As String
Private mOriginInvoiceId As Long
Private mUserId As Long

Private Sub Class_Initialize()
    mStatus = INVOICE_STATUS_ISSUED
    mInvoiceDate = Date
End Sub

Public Property Get Id() As Long
    Id = mIdentifier
End Property
Public Property Get InvoiceNumber() As String
    InvoiceNumber = mInvoiceNumber
End Property
Public Property Get MovementId() As Long
    MovementId = mMovementId
End Property
Public Property Get PartyId() As Long
    PartyId = mPartyId
End Property
Public Property Get InvoiceDate() As Date
    InvoiceDate = mInvoiceDate
End Property
Public Property Get Designation() As String
    Designation = mDesignation
End Property
Public Property Get NetAmount() As Currency
    NetAmount = mNetAmount
End Property
Public Property Get TaxRate() As Double
    TaxRate = mTaxRate
End Property
Public Property Get TaxAmount() As Currency
    TaxAmount = mTaxAmount
End Property
Public Property Get TotalAmount() As Currency
    TotalAmount = mTotalAmount
End Property
Public Property Get Status() As String
    Status = mStatus
End Property
Public Property Get OriginInvoiceId() As Long
    OriginInvoiceId = mOriginInvoiceId
End Property
Public Property Get UserId() As Long
    UserId = mUserId
End Property

' Prépare une facture émise : montants calculés à partir du TTC et du taux
Friend Sub Prepare(ByVal pInvoiceNumber As String, ByVal pMovementId As Long, ByVal pPartyId As Long, _
                   ByVal pInvoiceDate As Date, ByVal pDesignation As String, ByVal pTotalAmount As Currency, _
                   ByVal pTaxRate As Double, ByVal pUserId As Long)
    mInvoiceNumber = pInvoiceNumber
    mMovementId = pMovementId
    mPartyId = pPartyId
    mInvoiceDate = pInvoiceDate
    mDesignation = Left$(Trim$(pDesignation), MAX_DESIGNATION_LENGTH)
    mTotalAmount = pTotalAmount
    mTaxRate = pTaxRate
    SplitTaxIncluded pTotalAmount, mNetAmount, mTaxAmount, pTaxRate
    mStatus = INVOICE_STATUS_ISSUED
    mUserId = pUserId
End Sub

Friend Sub Hydrate(ByVal pId As Long, ByVal pInvoiceNumber As String, ByVal pMovementId As Long, ByVal pPartyId As Long, _
                   ByVal pInvoiceDate As Date, ByVal pDesignation As String, ByVal pNetAmount As Currency, _
                   ByVal pTaxRate As Double, ByVal pTaxAmount As Currency, ByVal pTotalAmount As Currency, _
                   ByVal pStatus As String, ByVal pOriginInvoiceId As Variant, ByVal pUserId As Long)
    mIdentifier = pId
    mInvoiceNumber = pInvoiceNumber
    mMovementId = pMovementId
    mPartyId = pPartyId
    mInvoiceDate = pInvoiceDate
    mDesignation = pDesignation
    mNetAmount = pNetAmount
    mTaxRate = pTaxRate
    mTaxAmount = pTaxAmount
    mTotalAmount = pTotalAmount
    mStatus = pStatus
    If IsNull(pOriginInvoiceId) Then mOriginInvoiceId = 0 Else mOriginInvoiceId = CLng(pOriginInvoiceId)
    mUserId = pUserId
End Sub

Friend Sub SetId(ByVal pId As Long)
    mIdentifier = pId
End Sub

' EMISE -> ANNULEE (le service vérifie le droit)
Public Sub Cancel()
    If mStatus <> INVOICE_STATUS_ISSUED Then
        Err.Raise ERR_INVOICE_INVALID, "Invoice.Cancel", "Cette facture est déjà annulée."
    End If
    mStatus = INVOICE_STATUS_CANCELLED
End Sub

Public Function ValidationErrors() As Collection
    Dim e As Collection

    Set e = New Collection
    If Len(mInvoiceNumber) = 0 Then e.Add "Le numéro de facture est obligatoire."
    If mMovementId <= 0 Then e.Add "La vente facturée est obligatoire."
    If mPartyId <= 0 Then e.Add "Le client est obligatoire."
    If Len(mDesignation) = 0 Then e.Add "La désignation est obligatoire."
    If mTotalAmount <= 0 Then e.Add "Le montant TTC doit être supérieur à zéro."
    If mNetAmount + mTaxAmount <> mTotalAmount Then e.Add "HT + taxe doit être égal au TTC."
    If mTaxRate < 0 Then e.Add "Le taux de taxe ne peut pas être négatif."
    If mUserId <= 0 Then e.Add "L'utilisateur qui émet la facture est inconnu."
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
    Err.Raise ERR_INVOICE_INVALID, "Invoice.Validate", msg
End Sub
