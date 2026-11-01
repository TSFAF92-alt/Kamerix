VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IContract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Interface IContract : ce que toute vente sait dire sur son règlement,
' quel que soit son mode (comptant : CashSale ; échelonnée : phase 6).
' Une classe qui écrit « Implements IContract » doit fournir chacun de ces
' membres sous le nom IContract_<membre>. Le code qui manipule un IContract
' fonctionne alors avec toutes les sortes de ventes (polymorphisme).
' Ce module ne contient aucun code : seulement les signatures.

' Mouvement de vente concerné
Public Property Get MovementId() As Long
End Property

' Prix convenu, TTC
Public Property Get Price() As Currency
End Property

' Total encaissé, remboursements déduits
Public Property Get AmountPaid() As Currency
End Property

' Reste à payer
Public Property Get Balance() As Currency
End Property

' Vrai si le prix est entièrement payé
Public Property Get IsSettled() As Boolean
End Property

' Vrai si un paiement de ce montant est acceptable (vente en cours, montant > 0, au plus le solde)
Public Function CanReceive(ByVal pAmount As Currency) As Boolean
End Function
