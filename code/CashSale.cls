VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CashSale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IContract

' Classe CashSale : vente comptant, vue sous l'angle de son règlement.
' Elle enveloppe le mouvement de vente et le total déjà encaissé.
' Plusieurs paiements sont possibles jusqu'au solde ; pas d'échéancier.
' Créée par le service de vente (modKamerixSales.LoadSaleContract).

Private mMovement As Movement
Private mAmountPaid As Currency

' Initialisation par le service : mouvement de vente et total encaissé
Friend Sub Init(ByVal pMovement As Movement, ByVal pAmountPaid As Currency)
    Set mMovement = pMovement
    mAmountPaid = pAmountPaid
End Sub

' Mouvement de vente enveloppé
Public Property Get SaleMovement() As Movement
    Set SaleMovement = mMovement
End Property

' ---------- Membres publics (mêmes réponses que l'interface) ----------

Public Property Get Price() As Currency
    Price = mMovement.Price
End Property

Public Property Get AmountPaid() As Currency
    AmountPaid = mAmountPaid
End Property

Public Property Get Balance() As Currency
    Balance = mMovement.Price - mAmountPaid
End Property

Public Property Get IsSettled() As Boolean
    IsSettled = (Me.Balance <= 0)
End Property

Public Function CanReceive(ByVal pAmount As Currency) As Boolean
    CanReceive = (mMovement.Status = MOVEMENT_STATUS_IN_PROGRESS) And (pAmount > 0) And (pAmount <= Me.Balance)
End Function

' ---------- Implémentation de IContract ----------

Private Property Get IContract_MovementId() As Long
    IContract_MovementId = mMovement.Id
End Property

Private Property Get IContract_Price() As Currency
    IContract_Price = Price
End Property

Private Property Get IContract_AmountPaid() As Currency
    IContract_AmountPaid = AmountPaid
End Property

Private Property Get IContract_Balance() As Currency
    IContract_Balance = Balance
End Property

Private Property Get IContract_IsSettled() As Boolean
    IContract_IsSettled = IsSettled
End Property

Private Function IContract_CanReceive(ByVal pAmount As Currency) As Boolean
    IContract_CanReceive = CanReceive(pAmount)
End Function
