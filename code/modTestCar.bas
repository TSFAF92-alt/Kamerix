Attribute VB_Name = "modTestCar"
Option Explicit
Option Private Module

Private mPassed As Long
Private mFailed As Long

Private Sub Check(ByVal pName As String, ByVal pCondition As Boolean)
    If pCondition Then
        mPassed = mPassed + 1
        Debug.Print "OK   - " & pName
    Else
        mFailed = mFailed + 1
        Debug.Print "ECHEC - " & pName
    End If
End Sub

Private Function NewCar(ByVal pStatus As String) As Car
    Dim v As Car
    Set v = New Car
    v.Hydrate 1, pStatus
    Set NewCar = v
End Function

Public Sub TestCar()
    Dim v As Car

    mPassed = 0
    mFailed = 0
    Debug.Print "=== Tests classe Car ==="

    ' 1. Parcours avec réservation jusqu'à la vente définitive
    Set v = NewCar(CAR_STATUS_IN_STOCK)
    v.Reserve
    v.ConfirmSale
    v.MarkAsSold
    Check "Parcours réservation -> vendu", (v.Status = CAR_STATUS_SOLD)

    ' 2. Désistement ou expiration : retour en stock
    Set v = NewCar(CAR_STATUS_IN_STOCK)
    v.Reserve
    v.CancelReservation
    Check "Réservé -> en stock", (v.Status = CAR_STATUS_IN_STOCK)

    ' 3. Vente directe puis annulation avant le CV
    Set v = NewCar(CAR_STATUS_IN_STOCK)
    v.StartDirectSale
    v.CancelSale "Client injoignable"
    Check "Vente en cours -> en stock", (v.Status = CAR_STATUS_IN_STOCK)

    ' 4. Retrait du stock avec motif
    Set v = NewCar(CAR_STATUS_IN_STOCK)
    v.Withdraw "Épave"
    Check "En stock -> retiré avec motif", (v.Status = CAR_STATUS_WITHDRAWN And v.WithdrawalReason = "Épave")

    ' 5. Un véhicule vendu ne peut plus être réservé
    Set v = NewCar(CAR_STATUS_SOLD)
    On Error Resume Next
    v.Reserve
    Check "Refus : vendu -> réservé", (Err.Number = ERR_CAR_TRANSITION)
    Err.Clear
    On Error GoTo 0

    ' 6. Retrait sans motif refusé
    Set v = NewCar(CAR_STATUS_IN_STOCK)
    On Error Resume Next
    v.Withdraw "   "
    Check "Refus : retrait sans motif", (Err.Number = ERR_CAR_REASON)
    Err.Clear
    On Error GoTo 0

    ' 7. Vente directe impossible depuis un véhicule réservé
    Set v = NewCar(CAR_STATUS_RESERVED)
    On Error Resume Next
    v.StartDirectSale
    Check "Refus : vente directe d'un véhicule réservé", (Err.Number = ERR_CAR_TRANSITION)
    Err.Clear
    On Error GoTo 0

    ' 8. Pas de saut d'étape : en stock -> vendu interdit
    Set v = NewCar(CAR_STATUS_IN_STOCK)
    On Error Resume Next
    v.MarkAsSold
    Check "Refus : en stock -> vendu", (Err.Number = ERR_CAR_TRANSITION)
    Err.Clear
    On Error GoTo 0

    ' 9. Statut inconnu refusé au chargement
    Set v = New Car
    On Error Resume Next
    v.Hydrate 1, "INCONNU"
    Check "Refus : statut inconnu", (Err.Number = ERR_CAR_STATUS)
    Err.Clear
    On Error GoTo 0

    Debug.Print "=== " & mPassed & " réussis, " & mFailed & " échoués ==="
End Sub
