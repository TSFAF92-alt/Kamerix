Attribute VB_Name = "modTestSalesScreens"
Option Explicit
Option Private Module

'==============================================================================
' KAMERIX - Tests du lot 5.4c : donnees de l'ecran des ventes
'------------------------------------------------------------------------------
' A lancer (F5 sur TestSalesScreens) sur une base de developpement SANS utilisateur.
' Tout se passe dans une transaction annulee a la fin : la base reste intacte.
'==============================================================================

Private Const PWD_ADMIN As String = "Secret#2026"

Public Sub TestSalesScreens()
    Dim lngOk As Long
    Dim lngTotal As Long
    Dim strReport As String
    Dim objData As DataAccess
    Dim objSeller As Party
    Dim objBuyer As Party
    Dim vData As Variant
    Dim lngModel As Long
    Dim lngCarA As Long
    Dim lngCarB As Long
    Dim lngCarC As Long
    Dim lngSaleA As Long
    Dim lngSaleB As Long
    Dim lngRow As Long
    Dim lngErr As Long
    Dim strErrText As String
    Dim blnTransaction As Boolean

    On Error GoTo EH

    If CLng(QueryValue("SELECT COUNT(*) FROM [Utilisateur]")) > 0 Then
        MsgBox "Ce test demande une base sans utilisateur." & vbCrLf & _
               "Supprimez la base de developpement et recreez-la, puis relancez le test.", _
               vbExclamation, "Kamerix - tests de l'ecran des ventes"
        Exit Sub
    End If

    CheckEq "Libelle du statut de vente", SaleStatusLabel(MOVEMENT_STATUS_SETTLED), "Sold" & ChrW(233) & "e", lngOk, lngTotal, strReport
    CheckEq "Libelle du statut du CV", CvStatusLabel(CV_STATUS_TO_FILE), ChrW(192) & " d" & ChrW(233) & "poser", lngOk, lngTotal, strReport

    CloseSession
    SetIterationsForTests 1000
    BeginTransaction
    blnTransaction = True

    CreateFirstAdmin "admin", "Administrateur Test", PWD_ADMIN
    OpenSession "admin", PWD_ADMIN
    lngModel = AddModel(AddBrand("TEST Toyota"), "Corolla")
    Set objData = New DataAccess
    Set objSeller = New Party: objSeller.Name = "TEST Vendeur": objSeller.IsSeller = True: objData.SaveParty objSeller
    Set objBuyer = New Party: objBuyer.Name = "TEST Client": objBuyer.IsBuyer = True: objData.SaveParty objBuyer

    lngCarA = PurchasedCar(lngModel, "TESTECR-A", objSeller.Id, 3000000)
    lngCarB = PurchasedCar(lngModel, "TESTECR-B", objSeller.Id, 2000000)
    lngCarC = NewCarId(objData, lngModel, "TESTECR-C")          ' en stock, sans achat

    ' ---------- Voitures vendables ----------
    vData = RecordsetToColumns(SaleableCarList(), Array("VoitureID", "Libelle"))
    CheckEq "Voitures vendables : en stock ET achetees (2)", CountId(vData, lngCarA) + CountId(vData, lngCarB) + _
        CountId(vData, lngCarC), 2, lngOk, lngTotal, strReport
    CheckTrue "Libelle : marque, modele et chassis", InStr(LabelOf(vData, lngCarA), "TEST Toyota Corolla - TESTECR-A") > 0, _
        lngOk, lngTotal, strReport

    lngSaleA = StartCashSale(lngCarA, objBuyer.Id, 4000000)
    RecordSalePayment lngSaleA, 1500000, "Esp" & ChrW(232) & "ces"
    lngSaleB = StartCashSale(lngCarB, objBuyer.Id, 2500000)
    RecordSalePayment lngSaleB, 1000000, "Esp" & ChrW(232) & "ces"
    CancelSale lngSaleB, "Test", False, "Esp" & ChrW(232) & "ces"

    vData = RecordsetToColumns(SaleableCarList(), Array("VoitureID", "Libelle"))
    CheckTrue "Voiture en vente : plus vendable ; voiture remise en stock : de nouveau vendable", _
        CountId(vData, lngCarA) = 0 And CountId(vData, lngCarB) = 1, lngOk, lngTotal, strReport

    CheckEq "Libelle d'une vente", SaleLabel(lngSaleA), "TEST Client - TEST Toyota Corolla - TESTECR-A", lngOk, lngTotal, strReport

    ' ---------- Liste des ventes ----------
    vData = SaleRows(MOVEMENT_STATUS_IN_PROGRESS)
    lngRow = RowOf(vData, lngSaleA)
    CheckTrue "Ventes en cours : la vente A, pas la vente annulee", lngRow >= 0 And RowOf(vData, lngSaleB) < 0, _
        lngOk, lngTotal, strReport
    CheckTrue "Vente en cours : prix, encaisse et solde", vData(SALE_COL_PRICE, lngRow) = Format$(4000000, "#,##0") And _
        vData(SALE_COL_PAID, lngRow) = Format$(1500000, "#,##0") And vData(SALE_COL_BALANCE, lngRow) = Format$(2500000, "#,##0"), _
        lngOk, lngTotal, strReport
    CheckTrue "Vente en cours : statuts (vente et CV)", vData(SALE_COL_STATUS, lngRow) = "En cours" And _
        vData(SALE_COL_CV_CODE, lngRow) = CV_STATUS_TO_FILE, lngOk, lngTotal, strReport

    vData = SaleRows(MOVEMENT_STATUS_CANCELLED)
    lngRow = RowOf(vData, lngSaleB)
    CheckTrue "Vente annulee : encaisse net 0 (rembourse), solde vide", lngRow >= 0 And _
        vData(SALE_COL_PAID, lngRow) = Format$(0, "#,##0") And vData(SALE_COL_BALANCE, lngRow) = "", lngOk, lngTotal, strReport

    vData = SaleRows("")
    CheckTrue "Toutes les ventes : les deux", RowOf(vData, lngSaleA) >= 0 And RowOf(vData, lngSaleB) >= 0, lngOk, lngTotal, strReport
    CheckTrue "Aucune vente soldee : Empty", IsEmpty(SaleRows(MOVEMENT_STATUS_SETTLED)), lngOk, lngTotal, strReport

    RecordSalePayment lngSaleA, 2500000, "Mobile Money"
    vData = SaleRows(MOVEMENT_STATUS_SETTLED)
    lngRow = RowOf(vData, lngSaleA)
    CheckTrue "Apres le solde : vente soldee, solde 0", lngRow >= 0 And vData(SALE_COL_BALANCE, lngRow) = Format$(0, "#,##0") And _
        vData(SALE_COL_STATUS_CODE, lngRow) = MOVEMENT_STATUS_SETTLED, lngOk, lngTotal, strReport

    ' ---------- Fin : on annule tout ----------
    CloseSession
    RollbackTransaction
    blnTransaction = False
    SetIterationsForTests 0

    CheckEq "La base est restee intacte (aucun utilisateur)", CLng(QueryValue("SELECT COUNT(*) FROM [Utilisateur]")), 0, _
        lngOk, lngTotal, strReport

    Debug.Print lngOk & " tests reussis sur " & lngTotal & "." & vbCrLf & strReport   ' rapport complet (Ctrl+G)
    MsgBox lngOk & " tests reussis sur " & lngTotal & "." & vbCrLf & vbCrLf & strReport, _
        IIf(lngOk = lngTotal, vbInformation, vbExclamation), "Kamerix - tests de l'ecran des ventes"
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
           "Resultats deja obtenus :" & vbCrLf & strReport, vbCritical, "Kamerix - tests de l'ecran des ventes"
End Sub

' Ligne dont la colonne 0 (identifiant) vaut pId ; -1 si absente
Private Function RowOf(ByVal vData As Variant, ByVal pId As Long) As Long
    Dim i As Long

    RowOf = -1
    If IsEmpty(vData) Then Exit Function
    For i = 0 To UBound(vData, 2)
        If CLng(vData(0, i)) = pId Then
            RowOf = i
            Exit Function
        End If
    Next i
End Function

Private Function CountId(ByVal vData As Variant, ByVal pId As Long) As Long
    If RowOf(vData, pId) >= 0 Then CountId = 1
End Function

Private Function LabelOf(ByVal vData As Variant, ByVal pId As Long) As String
    If RowOf(vData, pId) >= 0 Then LabelOf = CStr(vData(1, RowOf(vData, pId)))
End Function

Private Function PurchasedCar(ByVal pModelId As Long, ByVal pChassis As String, ByVal pSellerId As Long, _
                              ByVal pPrice As Currency) As Long
    Dim objCar As Car

    Set objCar = New Car
    objCar.ModelId = pModelId
    objCar.Chassis = pChassis
    RecordPurchase objCar, pSellerId, pPrice
    PurchasedCar = objCar.Id
End Function

Private Function NewCarId(ByVal pData As DataAccess, ByVal pModelId As Long, ByVal pChassis As String) As Long
    Dim objCar As Car

    Set objCar = New Car
    objCar.ModelId = pModelId
    objCar.Chassis = pChassis
    pData.SaveCar objCar
    NewCarId = objCar.Id
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
