Attribute VB_Name = "modTestDataAccess"
Option Explicit
Option Private Module
'==============================================================================
' Tests de la classe DataAccess (Car et Party)
' Lancer TestDataAccess avec F5. Tout est écrit dans une transaction annulée
' à la fin : la base reste inchangée.
'==============================================================================

Private mlngPassed As Long
Private mlngFailed As Long

Private Sub Check(ByVal strName As String, ByVal blnCondition As Boolean)
    If blnCondition Then
        mlngPassed = mlngPassed + 1
        Debug.Print "OK   - " & strName
    Else
        mlngFailed = mlngFailed + 1
        Debug.Print "ECHEC - " & strName
    End If
End Sub

' Vrai si la valeur n'est pas Null et égale au nombre attendu
Private Function IsEqualTo(ByVal vValue As Variant, ByVal lngExpected As Long) As Boolean
    If Not IsNull(vValue) Then IsEqualTo = (CLng(vValue) = lngExpected)
End Function

Private Function NewCar(ByVal lngModel As Long, ByVal strChassis As String) As Car
    Dim v As Car
    Set v = New Car
    v.ModelId = lngModel
    v.Chassis = strChassis
    Set NewCar = v
End Function

Public Sub TestDataAccess()
    Dim da As DataAccess
    Dim v As Car
    Dim w As Car
    Dim t As Party
    Dim u As Party
    Dim lngBrand As Long
    Dim lngModel As Long
    Dim lngCars0 As Long
    Dim lngParties0 As Long
    Dim lngId As Long

    mlngPassed = 0
    mlngFailed = 0
    Debug.Print "=== Tests classe DataAccess ==="

    On Error GoTo EH

    Set da = New DataAccess
    lngCars0 = CLng(QueryValue("SELECT COUNT(*) FROM [Voiture]"))
    lngParties0 = CLng(QueryValue("SELECT COUNT(*) FROM [Tiers]"))

    BeginTransaction

    ' Les voitures exigent un modèle existant (liens entre tables actifs)
    lngBrand = ExecuteInsert("INSERT INTO [Marque] ([Nom]) VALUES (?)", "TEST_MARQUE")
    lngModel = ExecuteInsert("INSERT INTO [Modele] ([MarqueID], [Nom]) VALUES (?, ?)", lngBrand, "TEST_MODELE")

    ' --- Voiture : insertion complète puis relecture
    Set v = NewCar(lngModel, "jtdbt123-45")
    v.ModelYear = 2002
    v.Mileage = 185000
    v.Color = "Grise"
    v.PlateNumber = "lt 123 ab"
    da.SaveCar v
    Check "Voiture : identifiant attribué à l'insertion", (v.Id > 0)

    Set w = da.LoadCar(v.Id)
    Check "Voiture : relecture non vide", Not (w Is Nothing)
    If Not (w Is Nothing) Then
        Check "Voiture : champs relus identiques", _
            (w.ModelId = lngModel And w.Chassis = "JTDBT123-45" And w.Color = "Grise" And _
             w.PlateNumber = "LT 123 AB" And IsEqualTo(w.ModelYear, 2002) And IsEqualTo(w.Mileage, 185000))
        Check "Voiture : statut En stock après insertion", (w.Status = CAR_STATUS_IN_STOCK)
        Check "Voiture : aucun motif de retrait pour une voiture en stock", (w.WithdrawalReason = "")
    End If

    ' --- Voiture : champs facultatifs vides (Null dans des colonnes Long et texte)
    Set v = NewCar(lngModel, "SANSOPTION1")
    v.ModelYear = Null
    v.Mileage = Null
    da.SaveCar v
    Set w = da.LoadCar(v.Id)
    Check "Voiture : champs facultatifs vides relus comme vides", _
        (Not (w Is Nothing)) And IsNull(w.ModelYear) And IsNull(w.Mileage) And (w.Color = "")

    ' --- Voiture : mise à jour et changement de statut enregistrés
    Set v = da.LoadCar(v.Id)
    v.Color = "Bleue"
    v.Reserve
    da.SaveCar v
    Set w = da.LoadCar(v.Id)
    Check "Voiture : mise à jour et statut Réservé relus", (w.Color = "Bleue" And w.Status = CAR_STATUS_RESERVED)

    w.ConfirmSale
    w.MarkAsSold
    da.SaveCar w
    Set w = da.LoadCar(w.Id)
    Check "Voiture : parcours jusqu'à Vendu enregistré", (w.Status = CAR_STATUS_SOLD)

    ' --- Voiture : retrait avec motif enregistré (colonne MotifRetrait, schéma v4)
    Set v = NewCar(lngModel, "RETIREE0001")
    v.Withdraw "Épave irréparable"
    da.SaveCar v
    Set w = da.LoadCar(v.Id)
    Check "Voiture : retrait et motif relus", _
        (Not (w Is Nothing)) And (w.Status = CAR_STATUS_WITHDRAWN) And (w.WithdrawalReason = "Épave irréparable")

    ' --- Voiture : introuvable
    Check "Voiture : identifiant inconnu renvoie Nothing", (da.LoadCar(999999) Is Nothing)

    ' --- Tiers : insertion complète puis relecture
    Set t = New Party
    t.Name = "Jean Mballa"
    t.Phone = "+237 6 55 12 34 56"
    t.IdType = "CNI"
    t.IdNumber = "123456789"
    t.Address = "Yaoundé"
    t.IsBuyer = True
    da.SaveParty t
    Check "Tiers : identifiant attribué à l'insertion", (t.Id > 0)

    Set u = da.LoadParty(t.Id)
    Check "Tiers : champs relus identiques", _
        (Not (u Is Nothing)) And (u.Name = "Jean Mballa") And (u.Phone = "+237 6 55 12 34 56") And _
        (u.IdType = "CNI") And (u.IdNumber = "123456789") And (u.Address = "Yaoundé") And _
        u.IsBuyer And Not u.IsSeller

    ' --- Tiers : champs facultatifs vides, vendeur seul
    Set t = New Party
    t.Name = "Awa"
    t.IsSeller = True
    da.SaveParty t
    Set u = da.LoadParty(t.Id)
    Check "Tiers : champs facultatifs vides, vendeur seul", _
        (Not (u Is Nothing)) And (u.Phone = "") And (u.IdType = "") And u.IsSeller And Not u.IsBuyer

    ' --- Tiers : mise à jour
    u.Name = "Awa Ndongo"
    u.IsBuyer = True
    da.SaveParty u
    Set u = da.LoadParty(u.Id)
    Check "Tiers : mise à jour relue", (u.Name = "Awa Ndongo" And u.IsBuyer And u.IsSeller)

    Check "Tiers : identifiant inconnu renvoie Nothing", (da.LoadParty(999999) Is Nothing)

    ' --- Erreurs attendues
    On Error Resume Next

    Set v = NewCar(lngModel, "JTDBT123-45")
    da.SaveCar v
    Check "Refus : châssis déjà utilisé", (Err.Number = ERR_DATA_DUPLICATE)
    Err.Clear

    Set v = NewCar(lngModel, "")
    da.SaveCar v
    Check "Refus : voiture invalide (châssis vide)", (Err.Number = ERR_CAR_DATA)
    Err.Clear

    Set v = New Car
    v.Hydrate 999999, CAR_STATUS_IN_STOCK
    v.HydrateData lngModel, Null, Null, "INCONNUE1", Null, Null
    da.SaveCar v
    Check "Refus : mise à jour d'une voiture inexistante", (Err.Number = ERR_DATA_NOT_FOUND)
    Err.Clear

    Set t = New Party
    t.Name = "Sans rôle"
    da.SaveParty t
    Check "Refus : tiers invalide (aucun rôle)", (Err.Number = ERR_PARTY_INVALID)
    Err.Clear

    On Error GoTo EH

    ' --- Annulation : la base revient à son état initial
    RollbackTransaction
    Check "Rollback : table Voiture inchangée", (CLng(QueryValue("SELECT COUNT(*) FROM [Voiture]")) = lngCars0)
    Check "Rollback : table Tiers inchangée", (CLng(QueryValue("SELECT COUNT(*) FROM [Tiers]")) = lngParties0)

    Debug.Print "=== " & mlngPassed & " réussis, " & mlngFailed & " échoués ==="
    Exit Sub

EH:
    Debug.Print "ERREUR INATTENDUE (" & Err.Number & ") : " & Err.Description
    mlngFailed = mlngFailed + 1
    RollbackTransaction
    Debug.Print "=== " & mlngPassed & " réussis, " & mlngFailed & " échoués ==="
End Sub
