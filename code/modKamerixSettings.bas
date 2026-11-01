Attribute VB_Name = "modKamerixSettings"
Option Explicit
Option Private Module

'==============================================================================
' KAMERIX - Parametres de gestion (table Parametre) et calcul des montants
'------------------------------------------------------------------------------
' Les valeurs viennent de la base, jamais du code : TAUX_TAXE, MODES_PAIEMENT,
' DEVISE, DECIMALES_DEVISE. Un parametre absent ou invalide leve ERR_SETTING.
' Prix      : les prix convenus sont TTC (taxe comprise). La facture en deduit
'             le HT et la taxe : HT = arrondi(TTC / (1 + taux)), taxe = TTC - HT,
'             de sorte que HT + taxe = TTC exactement.
' Arrondi   : au plus proche, la moitie vers le haut (et non l'arrondi bancaire
'             de la fonction Round de VBA), au nombre de decimales de la devise.
' Dependances : modKamerixSql, modKamerixConst.
'==============================================================================

Public Const PARAM_TAX_RATE As String = "TAUX_TAXE"
Public Const PARAM_PAYMENT_MODES As String = "MODES_PAIEMENT"
Public Const PARAM_CURRENCY As String = "DEVISE"
Public Const PARAM_CURRENCY_DECIMALS As String = "DECIMALES_DEVISE"

'==============================================================================
' PARAMETRES
'==============================================================================

' Taux de taxe en pourcentage (19.25 pour 19,25 %). Point ou virgule acceptes dans la base.
Public Function TaxRate() As Double
    Dim strValue As String
    Dim dblRate As Double

    strValue = Replace(RequiredSetting(PARAM_TAX_RATE), ",", ".")
    If Not IsNumericText(strValue) Then RaiseInvalid PARAM_TAX_RATE, strValue
    dblRate = Val(strValue)
    If dblRate < 0 Or dblRate >= 100 Then RaiseInvalid PARAM_TAX_RATE, strValue
    TaxRate = dblRate
End Function

' Modes de paiement proposes (tableau de textes, indices a partir de 0)
Public Function PaymentModes() As Variant
    Dim vParts As Variant
    Dim vModes() As String
    Dim lngCount As Long
    Dim i As Long

    vParts = Split(RequiredSetting(PARAM_PAYMENT_MODES), ";")
    ReDim vModes(0 To UBound(vParts))
    For i = LBound(vParts) To UBound(vParts)
        If Len(Trim$(vParts(i))) > 0 Then
            vModes(lngCount) = Trim$(vParts(i))
            lngCount = lngCount + 1
        End If
    Next i
    If lngCount = 0 Then RaiseInvalid PARAM_PAYMENT_MODES, "(vide)"

    ReDim Preserve vModes(0 To lngCount - 1)
    PaymentModes = vModes
End Function

' Vrai si le mode fait partie des modes de paiement (casse ignoree)
Public Function IsValidPaymentMode(ByVal pMode As String) As Boolean
    Dim vModes As Variant
    Dim i As Long

    vModes = PaymentModes()
    For i = LBound(vModes) To UBound(vModes)
        If StrComp(Trim$(pMode), vModes(i), vbTextCompare) = 0 Then
            IsValidPaymentMode = True
            Exit Function
        End If
    Next i
End Function

' Code de la devise affiche sur les documents (FCFA)
Public Function CurrencyCode() As String
    CurrencyCode = RequiredSetting(PARAM_CURRENCY)
End Function

' Nombre de decimales des montants (0 pour le FCFA)
Public Function CurrencyDecimals() As Long
    Dim strValue As String

    strValue = RequiredSetting(PARAM_CURRENCY_DECIMALS)
    If Not IsNumericText(strValue) Then RaiseInvalid PARAM_CURRENCY_DECIMALS, strValue
    If Val(strValue) < 0 Or Val(strValue) > 4 Or Val(strValue) <> Int(Val(strValue)) Then
        RaiseInvalid PARAM_CURRENCY_DECIMALS, strValue
    End If
    CurrencyDecimals = CLng(Val(strValue))
End Function

'==============================================================================
' MONTANTS
'==============================================================================

' Arrondi d'un montant aux decimales de la devise, la moitie vers le haut
' (en valeur absolue : -2,5 -> -3)
Public Function RoundAmount(ByVal pAmount As Variant) As Currency
    Dim decFactor As Variant
    Dim decValue As Variant

    decFactor = CDec(10 ^ CurrencyDecimals())
    decValue = CDec(pAmount)
    RoundAmount = CCur(Sgn(decValue) * Int(Abs(decValue) * decFactor + CDec(0.5)) / decFactor)
End Function

' Decompose un montant TTC : pNet (HT) et pTax (taxe), avec pNet + pTax = pTotal.
' pRate : taux en pourcentage ; par defaut, le taux du parametre TAUX_TAXE.
Public Sub SplitTaxIncluded(ByVal pTotal As Currency, ByRef pNet As Currency, ByRef pTax As Currency, _
                            Optional ByVal pRate As Variant)
    Dim dblRate As Double

    If IsMissing(pRate) Then dblRate = TaxRate() Else dblRate = CDbl(pRate)
    pNet = RoundAmount(CDec(pTotal) / CDec(1 + dblRate / 100))
    pTax = pTotal - pNet
End Sub

'==============================================================================
' PROCEDURES INTERNES
'==============================================================================

' Valeur texte d'un parametre ; erreur s'il est absent ou vide
Private Function RequiredSetting(ByVal pKey As String) As String
    Dim vValue As Variant

    vValue = QueryValue("SELECT [Valeur] FROM [Parametre] WHERE [Cle] = ?", pKey)
    If IsNull(vValue) Then
        Err.Raise ERR_SETTING, "modKamerixSettings", "Param" & ChrW(232) & "tre absent de la base : " & pKey & "."
    End If
    If Len(Trim$(CStr(vValue))) = 0 Then
        Err.Raise ERR_SETTING, "modKamerixSettings", "Param" & ChrW(232) & "tre vide dans la base : " & pKey & "."
    End If
    RequiredSetting = Trim$(CStr(vValue))
End Function

' Nombre ecrit avec un point decimal, independamment des reglages regionaux
Private Function IsNumericText(ByVal pText As String) As Boolean
    IsNumericText = (Len(pText) > 0) And Not (pText Like "*[!0-9.]*") And _
                    (Len(pText) - Len(Replace(pText, ".", "")) <= 1) And (pText <> ".")
End Function

Private Sub RaiseInvalid(ByVal pKey As String, ByVal pValue As String)
    Err.Raise ERR_SETTING, "modKamerixSettings", _
        "Param" & ChrW(232) & "tre invalide dans la base : " & pKey & " = " & pValue & "."
End Sub
