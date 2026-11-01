VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GeneralSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Classe GeneralSettings : paramètres de vente et de sécurité modifiables dans
' l'écran PARAMETRES (onglets Vente et Sécurité).
'   Vente    : TAUX_TAXE, MODES_PAIEMENT, DEVISE, DECIMALES_DEVISE
'   Sécurité : MAX_ECHECS_CONNEXION, DUREE_BLOCAGE_MINUTES
' Modes de paiement : saisis un par ligne ou séparés par ; ; stockés séparés par ;
' (20 caractères au plus chacun : taille de la colonne Paiement.ModePaiement).
' Le verrouillage de la devise et des décimales dès qu'une vente existe est
' contrôlé par le service (modKamerixCompany.SaveAllSettings), qui connaît la base.

Private Const MAX_PAYMENT_MODE_LENGTH As Long = 20
Private Const MAX_CURRENCY_LENGTH As Long = 10
Private Const MAX_DECIMALS As Long = 4
Private Const MAX_FAILED_LOGINS As Long = 20
Private Const MAX_LOCKOUT_MINUTES As Long = 1440

Private mTaxRate As Double
Private mPaymentModesText As String
Private mCurrencyCode As String
Private mCurrencyDecimals As Long
Private mMaxFailedLogins As Long
Private mLockoutMinutes As Long

' Taux de TVA en pourcentage (19.25 pour 19,25 %)
Public Property Get TaxRate() As Double
    TaxRate = mTaxRate
End Property
Public Property Let TaxRate(ByVal pValue As Double)
    mTaxRate = pValue
End Property

' Modes de paiement séparés par ; (la saisie peut en mettre un par ligne)
Public Property Get PaymentModesText() As String
    PaymentModesText = mPaymentModesText
End Property
Public Property Let PaymentModesText(ByVal pValue As String)
    mPaymentModesText = NormalizeModes(pValue)
End Property

Public Property Get CurrencyCode() As String
    CurrencyCode = mCurrencyCode
End Property
Public Property Let CurrencyCode(ByVal pValue As String)
    mCurrencyCode = UCase$(Trim$(pValue))
End Property

Public Property Get CurrencyDecimals() As Long
    CurrencyDecimals = mCurrencyDecimals
End Property
Public Property Let CurrencyDecimals(ByVal pValue As Long)
    mCurrencyDecimals = pValue
End Property

' Échecs de connexion consécutifs avant blocage du compte
Public Property Get MaxFailedLogins() As Long
    MaxFailedLogins = mMaxFailedLogins
End Property
Public Property Let MaxFailedLogins(ByVal pValue As Long)
    mMaxFailedLogins = pValue
End Property

' Durée du blocage, en minutes
Public Property Get LockoutMinutes() As Long
    LockoutMinutes = mLockoutMinutes
End Property
Public Property Let LockoutMinutes(ByVal pValue As Long)
    mLockoutMinutes = pValue
End Property

' Modes de paiement en tableau (indices à partir de 0) ; Empty s'il n'y en a aucun
Public Function PaymentModeList() As Variant
    If Len(mPaymentModesText) > 0 Then PaymentModeList = Split(mPaymentModesText, ";")
End Function

' ---------- Validation ----------

' Erreurs de l'onglet Vente
Public Function SaleErrors() As Collection
    Dim e As Collection
    Dim vModes As Variant
    Dim i As Long
    Dim j As Long

    Set e = New Collection
    If mTaxRate < 0 Or mTaxRate >= 100 Then e.Add "Le taux de TVA doit être compris entre 0 et 100 (exclu)."

    vModes = PaymentModeList()
    If IsEmpty(vModes) Then
        e.Add "Indiquez au moins un mode de paiement."
    Else
        For i = LBound(vModes) To UBound(vModes)
            If Len(vModes(i)) > MAX_PAYMENT_MODE_LENGTH Then
                e.Add "Mode de paiement trop long (" & MAX_PAYMENT_MODE_LENGTH & " caractères au plus) : " & vModes(i) & "."
            End If
            For j = i + 1 To UBound(vModes)
                If StrComp(vModes(i), vModes(j), vbTextCompare) = 0 Then e.Add "Mode de paiement en double : " & vModes(i) & "."
            Next j
        Next i
    End If

    If Len(mCurrencyCode) = 0 Then
        e.Add "La devise est obligatoire."
    ElseIf Len(mCurrencyCode) > MAX_CURRENCY_LENGTH Or InStr(mCurrencyCode, """") > 0 Then
        e.Add "La devise est invalide (" & MAX_CURRENCY_LENGTH & " caractères au plus, sans guillemet)."
    End If
    If mCurrencyDecimals < 0 Or mCurrencyDecimals > MAX_DECIMALS Then
        e.Add "Le nombre de décimales doit être compris entre 0 et " & MAX_DECIMALS & "."
    End If
    Set SaleErrors = e
End Function

' Erreurs de l'onglet Sécurité
Public Function SecurityErrors() As Collection
    Dim e As Collection

    Set e = New Collection
    If mMaxFailedLogins < 1 Or mMaxFailedLogins > MAX_FAILED_LOGINS Then
        e.Add "Le nombre d'échecs avant blocage doit être compris entre 1 et " & MAX_FAILED_LOGINS & "."
    End If
    If mLockoutMinutes < 1 Or mLockoutMinutes > MAX_LOCKOUT_MINUTES Then
        e.Add "La durée du blocage doit être comprise entre 1 et " & MAX_LOCKOUT_MINUTES & " minutes."
    End If
    Set SecurityErrors = e
End Function

' Toutes les erreurs (vente puis sécurité)
Public Function ValidationErrors() As Collection
    Dim e As Collection
    Dim v As Variant

    Set e = SaleErrors()
    For Each v In SecurityErrors()
        e.Add v
    Next v
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
    Err.Raise ERR_SETTINGS_INVALID, "GeneralSettings.Validate", msg
End Sub

' "Espèces" + saut de ligne + " Mobile Money " + ligne vide -> "Espèces;Mobile Money"
Private Function NormalizeModes(ByVal pText As String) As String
    Dim vParts As Variant
    Dim strResult As String
    Dim i As Long

    pText = Replace(Replace(Replace(pText, vbCrLf, ";"), vbCr, ";"), vbLf, ";")
    vParts = Split(pText, ";")
    For i = LBound(vParts) To UBound(vParts)
        If Len(Trim$(vParts(i))) > 0 Then
            If Len(strResult) > 0 Then strResult = strResult & ";"
            strResult = strResult & Trim$(vParts(i))
        End If
    Next i
    NormalizeModes = strResult
End Function
