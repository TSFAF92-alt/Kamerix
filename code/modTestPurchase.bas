Attribute VB_Name = "modTestPurchase"
Option Explicit
Option Private Module

'==============================================================================
' KAMERIX - Tests du lot 5.2 : enregistrement de l'achat d'une voiture
'------------------------------------------------------------------------------
' A lancer (F5 sur TestPurchase) sur une base de developpement SANS utilisateur.
' Tout se passe dans une transaction annulee a la fin : la base reste intacte.
'==============================================================================

Private Const PWD_ADMIN As String = "Secret#2026"
Private Const PWD_SELLER As String = "Vendeur#2026"

Public Sub TestPurchase()
    Dim lngOk As Long
    Dim lngTotal As Long
    Dim strReport As String
    Dim objData As DataAccess
    Dim objSeller As Party
    Dim objBuyerOnly As Party
    Dim objInter As Intermediary
    Dim objCar As Car
    Dim objCar2 As Car
    Dim objCar3 As Car
    Dim objMove As Movement
    Dim lngAdmin As Long
    Dim lngSalesman As Long
    Dim lngModel As Long
    Dim lngCity As Long
    Dim lngMove As Long
    Dim lngCategory As Long
    Dim lngMovesBefore As Long
    Dim lngCarsBefore As Long
    Dim lngErr As Long
    Dim strErrText As String
    Dim blnTransaction As Boolean

    On Error GoTo EH

    If CLng(QueryValue("SELECT COUNT(*) FROM [Utilisateur]")) > 0 Then
        MsgBox "Ce test demande une base sans utilisateur." & vbCrLf & _
               "Supprimez la base de developpement et recreez-la, puis relancez le test.", _
               vbExclamation, "Kamerix - tests de l'achat"
        Exit Sub
    End If

    lngMovesBefore = CLng(QueryValue("SELECT COUNT(*) FROM [Mouvement]"))
    lngCarsBefore = CLng(QueryValue("SELECT COUNT(*) FROM [Voiture]"))

    CloseSession
    SetIterationsForTests 1000
    BeginTransaction
    blnTransaction = True

    ' ---------- Preparation : comptes, referentiel, tiers ----------
    lngAdmin = CreateFirstAdmin("admin", "Administrateur Test", PWD_ADMIN)
    lngSalesman = CreateUser("vendeur1", "Vendeur Test", GetRoleIdByName("Vendeur"), PWD_SELLER)
    ChangePassword lngSalesman, PWD_SELLER, PWD_SELLER & "x"

    lngModel = AddModel(AddBrand("TEST Toyota"), "Corolla")
    lngCity = AddPlace("TEST Douala", PLACE_TYPE_CITY)

    Set objData = New DataAccess
    Set objSeller = NewParty("TEST Vendeur", False, True)
    objData.SaveParty objSeller
    Set objBuyerOnly = NewParty("TEST Acheteur", True, False)
    objData.SaveParty objBuyerOnly
    Set objInter = New Intermediary
    objInter.Name = "TEST Demarcheur"
    objInter.CityId = lngCity
    objData.SaveIntermediary objInter

    Set objCar = NewCar(lngModel, "TESTACH-001")

    ' ---------- Droits ----------
    OpenSession "vendeur1", PWD_SELLER & "x"
    CheckEq "Vendeur : achat refuse (pas d'ecriture sur Voitures)", _
        PurchaseError(objCar, objSeller.Id, 3500000), ERR_SESSION_DENIED, lngOk, lngTotal, strReport
    OpenSession "admin", PWD_ADMIN

    ' ---------- Regles ----------
    CheckEq "Tiers qui n'est pas vendeur refuse", PurchaseError(objCar, objBuyerOnly.Id, 3500000), ERR_PURCHASE, _
        lngOk, lngTotal, strReport
    CheckEq "Vendeur inconnu refuse", PurchaseError(objCar, 999999, 3500000), ERR_PURCHASE, lngOk, lngTotal, strReport
    CheckEq "Prix nul refuse", PurchaseError(objCar, objSeller.Id, 0), ERR_MOVEMENT_INVALID, lngOk, lngTotal, strReport
    CheckTrue "Echec : la voiture n'est pas enregistree (tout ou rien)", _
        objCar.Id = 0 And ChassisCount("TESTACH-001") = 0, lngOk, lngTotal, strReport
    CheckEq "Commission sans intermediaire refusee", PurchaseError(objCar, objSeller.Id, 3500000, 100000), _
        ERR_MOVEMENT_INVALID, lngOk, lngTotal, strReport
    CheckEq "Commission superieure au prix refusee", PurchaseError(objCar, objSeller.Id, 100000, 200000, objInter.Id), _
        ERR_MOVEMENT_INVALID, lngOk, lngTotal, strReport
    CheckEq "Intermediaire inconnu refuse", PurchaseError(objCar, objSeller.Id, 3500000, 0, 999999), ERR_PURCHASE, _
        lngOk, lngTotal, strReport
    CheckEq "Date dans le futur refusee", PurchaseError(objCar, objSeller.Id, 3500000, 0, 0, Date + 10), _
        ERR_MOVEMENT_INVALID, lngOk, lngTotal, strReport

    ' ---------- Achat valide ----------
    lngMove = RecordPurchase(objCar, objSeller.Id, 3500000, 150000, objInter.Id, lngCity, DateSerial(2026, 10, 15))
    CheckTrue "Achat enregistre : mouvement et voiture crees", lngMove > 0 And objCar.Id > 0, lngOk, lngTotal, strReport
    CheckEq "Voiture en stock apres l'achat", _
        QueryValue("SELECT [Statut] FROM [Voiture] WHERE [VoitureID] = ?", objCar.Id), CAR_STATUS_IN_STOCK, _
        lngOk, lngTotal, strReport

    Set objMove = objData.LoadMovement(lngMove)
    CheckTrue "Mouvement : achat comptant solde", objMove.MovementType = MOVEMENT_TYPE_PURCHASE And _
        objMove.SettlementMode = SETTLEMENT_CASH And objMove.Status = MOVEMENT_STATUS_SETTLED, lngOk, lngTotal, strReport
    CheckTrue "Mouvement : voiture et vendeur", objMove.CarId = objCar.Id And objMove.PartyId = objSeller.Id, _
        lngOk, lngTotal, strReport
    CheckTrue "Mouvement : prix et commission", objMove.Price = 3500000 And objMove.Commission = 150000, _
        lngOk, lngTotal, strReport
    CheckTrue "Mouvement : intermediaire et lieu", objMove.IntermediaryId = objInter.Id And objMove.PlaceId = lngCity, _
        lngOk, lngTotal, strReport
    CheckTrue "Mouvement : date de l'achat", objMove.MovementDate = DateSerial(2026, 10, 15), lngOk, lngTotal, strReport
    CheckEq "Mouvement : utilisateur connecte inscrit", objMove.UserId, lngAdmin, lngOk, lngTotal, strReport
    CheckEq "PurchaseOfCar retrouve l'achat", PurchaseOfCar(objCar.Id).Id, lngMove, lngOk, lngTotal, strReport
    CheckEq "Deuxieme achat de la meme voiture refuse", PurchaseError(objCar, objSeller.Id, 3500000), ERR_PURCHASE, _
        lngOk, lngTotal, strReport

    ' ---------- Voiture deja saisie, sans intermediaire ----------
    Set objCar2 = NewCar(lngModel, "TESTACH-002")
    objData.SaveCar objCar2
    lngMove = RecordPurchase(objCar2, objSeller.Id, 2800000)
    CheckTrue "Achat d'une voiture deja saisie", lngMove > 0, lngOk, lngTotal, strReport
    CheckTrue "Sans intermediaire : colonnes vides dans la base", _
        IsNull(QueryValue("SELECT [IntermediaireID] FROM [Mouvement] WHERE [MouvementID] = ?", lngMove)) And _
        IsNull(QueryValue("SELECT [LieuID] FROM [Mouvement] WHERE [MouvementID] = ?", lngMove)), lngOk, lngTotal, strReport
    CheckTrue "Date par defaut : aujourd'hui", objData.LoadMovement(lngMove).MovementDate = Date, lngOk, lngTotal, strReport

    ' ---------- Voiture qui n'est pas en stock ----------
    Set objCar3 = NewCar(lngModel, "TESTACH-003")
    objCar3.Reserve
    CheckEq "Voiture reservee : achat refuse", PurchaseError(objCar3, objSeller.Id, 1000000), ERR_PURCHASE, _
        lngOk, lngTotal, strReport

    ' ---------- Cout de revient ----------
    CheckEq "Cout de revient sans charge = prix d'achat + commission", CarCostPrice(objCar.Id), 3650000, lngOk, lngTotal, strReport
    CheckEq "Cout de revient sans commission ni charge = prix d'achat", CarCostPrice(objCar2.Id), 2800000, lngOk, lngTotal, strReport
    lngCategory = CLng(QueryValue("SELECT [CategorieID] FROM [CategorieCharge] WHERE [Libelle] = ?", "Lavage"))
    ExecuteInsert "INSERT INTO [Charge] ([VoitureID], [CategorieID], [DateCharge], [Montant], [UtilisateurID]) " & _
                  "VALUES (?, ?, ?, ?, ?)", objCar.Id, lngCategory, Date, CCur(25000), lngAdmin
    ExecuteInsert "INSERT INTO [Charge] ([VoitureID], [CategorieID], [DateCharge], [Montant], [UtilisateurID]) " & _
                  "VALUES (?, ?, ?, ?, ?)", objCar.Id, lngCategory, Date, CCur(40000), lngAdmin
    CheckEq "Cout de revient = prix d'achat + commission + charges", CarCostPrice(objCar.Id), 3715000, lngOk, lngTotal, strReport
    CheckEq "Cout de revient d'une voiture sans achat refuse", CostError(999999), ERR_PURCHASE, lngOk, lngTotal, strReport

    ' ---------- Fin : on annule tout ----------
    CloseSession
    RollbackTransaction
    blnTransaction = False
    SetIterationsForTests 0

    CheckEq "La base est restee intacte (mouvements)", CLng(QueryValue("SELECT COUNT(*) FROM [Mouvement]")), lngMovesBefore, _
        lngOk, lngTotal, strReport
    CheckEq "La base est restee intacte (voitures)", CLng(QueryValue("SELECT COUNT(*) FROM [Voiture]")), lngCarsBefore, _
        lngOk, lngTotal, strReport
    CheckEq "La base est restee intacte (aucun utilisateur)", CLng(QueryValue("SELECT COUNT(*) FROM [Utilisateur]")), 0, _
        lngOk, lngTotal, strReport

    Debug.Print lngOk & " tests reussis sur " & lngTotal & "." & vbCrLf & strReport   ' rapport complet (Ctrl+G)
    MsgBox lngOk & " tests reussis sur " & lngTotal & "." & vbCrLf & vbCrLf & strReport, _
        IIf(lngOk = lngTotal, vbInformation, vbExclamation), "Kamerix - tests de l'achat"
    Exit Sub

EH:
    lngErr = Err.Number
    strErrText = Err.Description
    On Error Resume Next
    CloseSession
    If blnTransaction Then RollbackTransaction
    SetIterationsForTests 0
    On Error GoTo 0
    MsgBox "Echec (erreur " & lngErr & ") :" & vbCrLf & vbCrLf & strErrText & vbCrLf & vbCrLf & _
           "Resultats deja obtenus :" & vbCrLf & strReport, vbCritical, "Kamerix - tests de l'achat"
End Sub

' Numero d'erreur d'un achat (0 si accepte)
Private Function PurchaseError(ByVal pCar As Car, ByVal pSellerId As Long, ByVal pPrice As Currency, _
                               Optional ByVal pCommission As Currency = 0, Optional ByVal pIntermediaryId As Long = 0, _
                               Optional ByVal pDate As Variant) As Long
    On Error Resume Next
    If IsMissing(pDate) Then
        RecordPurchase pCar, pSellerId, pPrice, pCommission, pIntermediaryId
    Else
        RecordPurchase pCar, pSellerId, pPrice, pCommission, pIntermediaryId, 0, pDate
    End If
    PurchaseError = Err.Number
    Err.Clear
End Function

Private Function CostError(ByVal pCarId As Long) As Long
    On Error Resume Next
    CarCostPrice pCarId
    CostError = Err.Number
    Err.Clear
End Function

Private Function NewParty(ByVal pName As String, ByVal pIsBuyer As Boolean, ByVal pIsSeller As Boolean) As Party
    Set NewParty = New Party
    NewParty.Name = pName
    NewParty.IsBuyer = pIsBuyer
    NewParty.IsSeller = pIsSeller
End Function

Private Function NewCar(ByVal pModelId As Long, ByVal pChassis As String) As Car
    Set NewCar = New Car
    NewCar.ModelId = pModelId
    NewCar.Chassis = pChassis
    NewCar.ModelYear = 2015
End Function

Private Function ChassisCount(ByVal pChassis As String) As Long
    ChassisCount = CLng(QueryValue("SELECT COUNT(*) FROM [Voiture] WHERE [Chassis] = ?", pChassis))
End Function

' Comparaison de deux valeurs (converties en texte) ; affiche obtenu et attendu en cas d'echec
Private Sub CheckEq(ByVal strLabel As String, ByVal vActual As Variant, ByVal vExpected As Variant, _
                    ByRef lngOk As Long, ByRef lngTotal As Long, ByRef strReport As String)
    lngTotal = lngTotal + 1
    If IsNull(vActual) Then vActual = "(Null)"
    If CStr(vActual) = CStr(vExpected) Then
        lngOk = lngOk + 1
        strReport = strReport & "OK     " & strLabel & vbCrLf
    Else
        strReport = strReport & "ECHEC  " & strLabel & " (obtenu " & CStr(vActual) & ", attendu " & CStr(vExpected) & ")" & vbCrLf
    End If
End Sub

' Verification d'une condition vraie (independant de la langue de Windows/Office)
Private Sub CheckTrue(ByVal strLabel As String, ByVal blnCondition As Boolean, _
                      ByRef lngOk As Long, ByRef lngTotal As Long, ByRef strReport As String)
    lngTotal = lngTotal + 1
    If blnCondition Then
        lngOk = lngOk + 1
        strReport = strReport & "OK     " & strLabel & vbCrLf
    Else
        strReport = strReport & "ECHEC  " & strLabel & vbCrLf
    End If
End Sub
