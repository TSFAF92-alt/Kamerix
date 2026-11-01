Attribute VB_Name = "modTestPartyCarData"
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

Private Function NewValidParty() As Party
    Dim t As Party
    Set t = New Party
    t.Name = "  Jean Mballa "
    t.Phone = "+237 6 55 12 34 56"
    t.IdType = "CNI"
    t.IdNumber = "123456789"
    t.IsBuyer = True
    Set NewValidParty = t
End Function

Private Function NewValidCar() As Car
    Dim v As Car
    Set v = New Car
    v.ModelId = 1
    v.Chassis = "jtdbt123-45"
    v.ModelYear = 2002
    v.Mileage = 185000
    Set NewValidCar = v
End Function

Public Sub TestParty()
    Dim t As Party

    mPassed = 0
    mFailed = 0
    Debug.Print "=== Tests classe Party ==="

    Set t = NewValidParty()
    Check "Tiers valide : aucune erreur", (t.ValidationErrors.Count = 0)
    Check "Le nom est nettoyé des espaces", (t.Name = "Jean Mballa")

    Set t = NewValidParty()
    t.Name = "   "
    Check "Refus : nom vide", (t.ValidationErrors.Count = 1)

    Set t = NewValidParty()
    t.IsBuyer = False
    t.IsSeller = False
    Check "Refus : aucun rôle", (t.ValidationErrors.Count = 1)

    Set t = NewValidParty()
    t.IsBuyer = False
    t.IsSeller = True
    Check "Vendeur seul accepté", (t.ValidationErrors.Count = 0)

    Set t = NewValidParty()
    t.Phone = "abc"
    Check "Refus : téléphone avec des lettres", (t.ValidationErrors.Count = 1)

    Set t = NewValidParty()
    t.Phone = "123"
    Check "Refus : téléphone trop court", (t.ValidationErrors.Count = 1)

    Set t = NewValidParty()
    t.Phone = ""
    Check "Téléphone facultatif", (t.ValidationErrors.Count = 0)

    Set t = NewValidParty()
    t.IdNumber = ""
    Check "Refus : type de pièce sans numéro", (t.ValidationErrors.Count = 1)

    Set t = NewValidParty()
    t.Name = ""
    t.IsBuyer = False
    On Error Resume Next
    t.Validate
    Check "Validate regroupe les erreurs dans une seule erreur", _
        (Err.Number = ERR_PARTY_INVALID And InStr(Err.Description, "nom") > 0 And InStr(Err.Description, "acheteur") > 0)
    Err.Clear
    On Error GoTo 0

    Set t = New Party
    t.Hydrate 7, "Awa", Null, Null, Null, Null, Null, True
    Check "HydrateData : colonnes Null lues comme vides", _
        (t.Id = 7 And t.Phone = "" And t.Address = "" And t.IsBuyer = False And t.IsSeller = True)

    Debug.Print "=== " & mPassed & " réussis, " & mFailed & " échoués ==="
End Sub

Public Sub TestCarData()
    Dim v As Car

    mPassed = 0
    mFailed = 0
    Debug.Print "=== Tests données de la classe Car ==="

    Set v = New Car
    Check "Voiture neuve : statut En stock", (v.Status = CAR_STATUS_IN_STOCK)

    Set v = NewValidCar()
    Check "Voiture valide : aucune erreur", (v.ValidationErrors.Count = 0)
    Check "Le châssis est mis en majuscules", (v.Chassis = "JTDBT123-45")

    Set v = NewValidCar()
    v.Chassis = ""
    Check "Refus : châssis vide", (v.ValidationErrors.Count = 1)

    Set v = NewValidCar()
    v.Chassis = "AB CD"
    Check "Refus : châssis avec un espace", (v.ValidationErrors.Count = 1)

    Set v = NewValidCar()
    v.ModelId = 0
    Check "Refus : modèle absent", (v.ValidationErrors.Count = 1)

    Set v = NewValidCar()
    v.ModelYear = 1900
    Check "Refus : année trop ancienne", (v.ValidationErrors.Count = 1)

    Set v = NewValidCar()
    v.ModelYear = Null
    v.Mileage = Null
    Check "Année et kilométrage facultatifs", (v.ValidationErrors.Count = 0)

    Set v = NewValidCar()
    v.Mileage = -5
    Check "Refus : kilométrage négatif", (v.ValidationErrors.Count = 1)

    Set v = NewValidCar()
    v.ModelYear = "deux mille"
    Check "Refus : année non numérique", (v.ValidationErrors.Count = 1)

    Set v = NewValidCar()
    v.Chassis = ""
    On Error Resume Next
    v.Validate
    Check "Validate lève ERR_CAR_DATA", (Err.Number = ERR_CAR_DATA)
    Err.Clear
    On Error GoTo 0

    Set v = New Car
    v.Hydrate 1, CAR_STATUS_SOLD
    On Error Resume Next
    v.Color = "Rouge"
    Check "Refus : modifier une voiture vendue", (Err.Number = ERR_CAR_LOCKED)
    Err.Clear
    On Error GoTo 0

    Set v = New Car
    v.HydrateData 3, Null, Null, "abc123", Null, Null
    Check "HydrateData : Null accepté pour les champs facultatifs", _
        (v.ModelId = 3 And IsNull(v.ModelYear) And v.Color = "" And v.Chassis = "ABC123" And IsNull(v.Mileage))

    Debug.Print "=== " & mPassed & " réussis, " & mFailed & " échoués ==="
End Sub
