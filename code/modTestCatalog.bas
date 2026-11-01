Attribute VB_Name = "modTestCatalog"
Option Explicit
Option Private Module

'==============================================================================
' KAMERIX - Tests du lot 5.4b : services des ecrans tiers, intermediaires, voitures
'------------------------------------------------------------------------------
' A lancer (F5 sur TestCatalog) sur une base de developpement SANS utilisateur.
' Tout se passe dans une transaction annulee a la fin : la base reste intacte.
'==============================================================================

Private Const PWD_ADMIN As String = "Secret#2026"
Private Const PWD_OTHER As String = "Autre#20260"

Public Sub TestCatalog()
    Dim lngOk As Long
    Dim lngTotal As Long
    Dim strReport As String
    Dim curAmount As Currency
    Dim objData As DataAccess
    Dim objParty As Party
    Dim objInter As Intermediary
    Dim objCar As Car
    Dim objCar2 As Car
    Dim vData As Variant
    Dim lngSalesman As Long
    Dim lngAccountant As Long
    Dim lngBrand As Long
    Dim lngModel As Long
    Dim lngCity As Long
    Dim lngSeller As Long
    Dim lngErr As Long
    Dim strErrText As String
    Dim blnTransaction As Boolean

    On Error GoTo EH

    If CLng(QueryValue("SELECT COUNT(*) FROM [Utilisateur]")) > 0 Then
        MsgBox "Ce test demande une base sans utilisateur." & vbCrLf & _
               "Supprimez la base de developpement et recreez-la, puis relancez le test.", _
               vbExclamation, "Kamerix - tests des ecrans"
        Exit Sub
    End If

    ' ---------- Lecture des montants saisis ----------
    CheckTrue "Montant avec espaces : 3 500 000", TryParseAmount("3 500 000", curAmount) And curAmount = 3500000, _
        lngOk, lngTotal, strReport
    CheckTrue "Montant avec espaces insecables", TryParseAmount("1" & ChrW(160) & "250" & ChrW(160) & "000", curAmount) _
        And curAmount = 1250000, lngOk, lngTotal, strReport
    CheckTrue "Montant vide : 0", TryParseAmount("  ", curAmount) And curAmount = 0, lngOk, lngTotal, strReport
    CheckTrue "Texte refuse", Not TryParseAmount("abc", curAmount), lngOk, lngTotal, strReport
    CheckTrue "Montant negatif refuse", Not TryParseAmount("-5", curAmount), lngOk, lngTotal, strReport
    CheckEq "Libelle d'un statut", CarStatusLabel(CAR_STATUS_SALE_IN_PROGRESS), "Vente en cours", lngOk, lngTotal, strReport

    CloseSession
    SetIterationsForTests 1000
    BeginTransaction
    blnTransaction = True

    CreateFirstAdmin "admin", "Administrateur Test", PWD_ADMIN
    lngSalesman = CreateUser("vendeur1", "Vendeur Test", GetRoleIdByName("Vendeur"), "Provisoire#1")
    ChangePassword lngSalesman, "Provisoire#1", PWD_OTHER
    lngAccountant = CreateUser("compta1", "Comptable Test", GetRoleIdByName("Comptable"), "Provisoire#1")
    ChangePassword lngAccountant, "Provisoire#1", PWD_OTHER

    lngBrand = AddBrand("TEST Toyota")
    lngModel = AddModel(lngBrand, "Corolla")
    lngCity = AddPlace("TEST Douala", PLACE_TYPE_CITY)
    CheckEq "Marque d'un modele", BrandOfModel(lngModel), lngBrand, lngOk, lngTotal, strReport

    ' ---------- Tiers ----------
    OpenSession "compta1", PWD_OTHER
    Set objParty = NewParty("TEST Kouam Paul", True, False, "CNI", "123456")
    CheckEq "Comptable : enregistrement d'un tiers refuse (Tiers en lecture)", SaveError(objParty), ERR_SESSION_DENIED, _
        lngOk, lngTotal, strReport

    OpenSession "vendeur1", PWD_OTHER
    SavePartyChecked objParty
    CheckTrue "Vendeur : tiers enregistre (Tiers en ecriture)", objParty.Id > 0, lngOk, lngTotal, strReport
    Set objParty = NewParty("TEST Mballa Jean", False, True, "", "")
    SavePartyChecked objParty
    lngSeller = objParty.Id
    Set objParty = NewParty("TEST Ngo Marie", True, True, "Passeport", "P-77")
    SavePartyChecked objParty

    vData = RecordsetToColumns(PartyList("TEST"), Array("Nom", "Piece", "Roles"))
    CheckEq "Liste des tiers : 3 tiers de test", UBound(vData, 2) + 1, 3, lngOk, lngTotal, strReport
    CheckTrue "Liste triee par nom ; piece = type + numero", vData(0, 0) = "TEST Kouam Paul" And vData(1, 0) = "CNI 123456", _
        lngOk, lngTotal, strReport
    CheckTrue "Sans piece : colonne vide", vData(1, 1) = "", lngOk, lngTotal, strReport
    CheckEq "Role affiche : acheteur et vendeur", vData(2, 2), "Acheteur et vendeur", lngOk, lngTotal, strReport
    vData = RecordsetToColumns(PartyList("kouam"), Array("Nom"))
    CheckTrue "Recherche par partie du nom (casse ignoree)", UBound(vData, 2) = 0, lngOk, lngTotal, strReport
    vData = RecordsetToColumns(PartyList("TEST", PARTY_ROLE_BUYER), Array("Nom"))
    CheckEq "Filtre Acheteurs : 2 (dont acheteur et vendeur)", UBound(vData, 2) + 1, 2, lngOk, lngTotal, strReport
    vData = RecordsetToColumns(PartyList("TEST", PARTY_ROLE_SELLER), Array("Nom"))
    CheckEq "Filtre Vendeurs : 2", UBound(vData, 2) + 1, 2, lngOk, lngTotal, strReport

    Set objInter = New Intermediary
    objInter.Name = "TEST Demarcheur"
    objInter.CityId = lngCity
    SaveIntermediaryChecked objInter
    vData = RecordsetToColumns(IntermediaryList("demarch"), Array("Nom", "NomVille"))
    CheckTrue "Intermediaire enregistre et trouve par recherche, avec sa ville", _
        objInter.Id > 0 And vData(1, 0) = "TEST Douala", lngOk, lngTotal, strReport

    ' ---------- Voitures ----------
    CheckEq "Vendeur : retrait du stock refuse (Voitures en lecture)", WithdrawError(999999, "x"), ERR_SESSION_DENIED, _
        lngOk, lngTotal, strReport

    OpenSession "admin", PWD_ADMIN
    Set objData = New DataAccess
    Set objCar = NewCar(lngModel, "TESTCAT-001")
    RecordPurchase objCar, lngSeller, 3500000
    Set objCar2 = NewCar(lngModel, "TESTCAT-002")
    RecordPurchase objCar2, lngSeller, 2000000

    CheckEq "Nouvelle voiture sans achat refusee par SaveCarDetails", SaveCarError(NewCar(lngModel, "TESTCAT-003")), _
        ERR_PURCHASE, lngOk, lngTotal, strReport
    objCar2.Color = "Rouge"
    SaveCarDetails objCar2
    CheckEq "Fiche voiture : modification enregistree", objData.LoadCar(objCar2.Id).Color, "Rouge", lngOk, lngTotal, strReport

    CheckEq "Retrait sans motif refuse", WithdrawError(objCar2.Id, "  "), ERR_CAR_REASON, lngOk, lngTotal, strReport
    WithdrawCar objCar2.Id, "Accident, voiture irreparable"
    Set objCar2 = objData.LoadCar(objCar2.Id)
    CheckTrue "Voiture retiree avec son motif", objCar2.Status = CAR_STATUS_WITHDRAWN And _
        objCar2.WithdrawalReason = "Accident, voiture irreparable", lngOk, lngTotal, strReport
    CheckEq "Voiture retiree : nouveau retrait refuse", WithdrawError(objCar2.Id, "Encore"), ERR_CAR_TRANSITION, _
        lngOk, lngTotal, strReport
    CheckEq "Voiture retiree : donnees verrouillees", LockedError(objCar2), ERR_CAR_LOCKED, lngOk, lngTotal, strReport

    vData = RecordsetToColumns(CarList(""), Array("Chassis", "LibelleStatut", "PrixAchat", "Modele"))
    CheckEq "Liste des voitures : toutes (2)", CountChassis(vData, "TESTCAT-"), 2, lngOk, lngTotal, strReport
    CheckEq "Libelle du modele : marque + modele", vData(3, 0), "TEST Toyota Corolla", lngOk, lngTotal, strReport
    vData = RecordsetToColumns(CarList(CAR_STATUS_IN_STOCK), Array("Chassis", "LibelleStatut", "PrixAchat"))
    CheckTrue "Filtre En stock : la voiture non retiree, avec son prix d'achat", CountChassis(vData, "TESTCAT-") = 1 _
        And RowOf(vData, "TESTCAT-001") >= 0, lngOk, lngTotal, strReport
    CheckTrue "Statut et prix d'achat dans la liste", vData(1, RowOf(vData, "TESTCAT-001")) = "En stock" And _
        CCur(vData(2, RowOf(vData, "TESTCAT-001"))) = 3500000, lngOk, lngTotal, strReport
    vData = RecordsetToColumns(CarList(CAR_STATUS_WITHDRAWN), Array("Chassis", "LibelleStatut"))
    CheckEq "Filtre Retirees : libelle", vData(1, RowOf(vData, "TESTCAT-002")), "Retir" & ChrW(233) & "e", lngOk, lngTotal, strReport

    ' ---------- Fin : on annule tout ----------
    CloseSession
    RollbackTransaction
    blnTransaction = False
    SetIterationsForTests 0

    CheckEq "La base est restee intacte (aucun tiers de test)", _
        CLng(QueryValue("SELECT COUNT(*) FROM [Tiers] WHERE [Nom] LIKE ?", "TEST %")), 0, lngOk, lngTotal, strReport
    CheckEq "La base est restee intacte (aucun utilisateur)", CLng(QueryValue("SELECT COUNT(*) FROM [Utilisateur]")), 0, _
        lngOk, lngTotal, strReport

    Debug.Print lngOk & " tests reussis sur " & lngTotal & "." & vbCrLf & strReport   ' rapport complet (Ctrl+G)
    MsgBox lngOk & " tests reussis sur " & lngTotal & "." & vbCrLf & vbCrLf & strReport, _
        IIf(lngOk = lngTotal, vbInformation, vbExclamation), "Kamerix - tests des ecrans"
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
           "Resultats deja obtenus :" & vbCrLf & strReport, vbCritical, "Kamerix - tests des ecrans"
End Sub

Private Function SaveError(ByVal pParty As Party) As Long
    On Error Resume Next
    SavePartyChecked pParty
    SaveError = Err.Number
    Err.Clear
End Function

Private Function WithdrawError(ByVal pCarId As Long, ByVal pReason As String) As Long
    On Error Resume Next
    WithdrawCar pCarId, pReason
    WithdrawError = Err.Number
    Err.Clear
End Function

Private Function SaveCarError(ByVal pCar As Car) As Long
    On Error Resume Next
    SaveCarDetails pCar
    SaveCarError = Err.Number
    Err.Clear
End Function

Private Function LockedError(ByVal pCar As Car) As Long
    On Error Resume Next
    pCar.Color = "Bleu"
    LockedError = Err.Number
    Err.Clear
End Function

' Nombre de lignes dont le chassis (colonne 0) commence par pPrefix
Private Function CountChassis(ByVal vData As Variant, ByVal pPrefix As String) As Long
    Dim i As Long

    If IsEmpty(vData) Then Exit Function
    For i = 0 To UBound(vData, 2)
        If Left$(CStr(vData(0, i)), Len(pPrefix)) = pPrefix Then CountChassis = CountChassis + 1
    Next i
End Function

' Ligne dont le chassis (colonne 0) vaut pChassis ; -1 si absente
Private Function RowOf(ByVal vData As Variant, ByVal pChassis As String) As Long
    Dim i As Long

    RowOf = -1
    If IsEmpty(vData) Then Exit Function
    For i = 0 To UBound(vData, 2)
        If CStr(vData(0, i)) = pChassis Then
            RowOf = i
            Exit Function
        End If
    Next i
End Function

Private Function NewParty(ByVal pName As String, ByVal pIsBuyer As Boolean, ByVal pIsSeller As Boolean, _
                          ByVal pIdType As String, ByVal pIdNumber As String) As Party
    Set NewParty = New Party
    NewParty.Name = pName
    NewParty.IsBuyer = pIsBuyer
    NewParty.IsSeller = pIsSeller
    NewParty.IdType = pIdType
    NewParty.IdNumber = pIdNumber
End Function

Private Function NewCar(ByVal pModelId As Long, ByVal pChassis As String) As Car
    Set NewCar = New Car
    NewCar.ModelId = pModelId
    NewCar.Chassis = pChassis
    NewCar.ModelYear = 2016
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
