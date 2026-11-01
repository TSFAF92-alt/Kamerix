Attribute VB_Name = "modTestReference"
Option Explicit
Option Private Module

'==============================================================================
' KAMERIX - Tests du lot 5.1 : parametres, montants, numerotation, referentiel
'------------------------------------------------------------------------------
' A lancer (F5 sur TestReference) sur une base creee avec le schema du lot 5.1
' (parametres TAUX_TAXE, MODES_PAIEMENT, DEVISE, DECIMALES_DEVISE inscrits).
' Tout se passe dans une transaction annulee a la fin : la base reste intacte.
'==============================================================================

Public Sub TestReference()
    Dim lngOk As Long
    Dim lngTotal As Long
    Dim strReport As String
    Dim curNet As Currency
    Dim curTax As Currency
    Dim vModes As Variant
    Dim lngBrand As Long
    Dim lngBrand2 As Long
    Dim lngModel As Long
    Dim lngCity As Long
    Dim lngDistrict As Long
    Dim lngStation As Long
    Dim lngCounters As Long
    Dim lngPlaces As Long
    Dim objData As DataAccess
    Dim objInter As Intermediary
    Dim objRead As Intermediary
    Dim strNumber As String
    Dim lngErr As Long
    Dim strErrText As String
    Dim blnTransaction As Boolean

    On Error GoTo EH

    ' ---------- Parametres ----------
    CheckEq "Taux de taxe : 19,25 %", TaxRate(), 19.25, lngOk, lngTotal, strReport
    vModes = PaymentModes()
    CheckEq "Modes de paiement : 3", UBound(vModes) - LBound(vModes) + 1, 3, lngOk, lngTotal, strReport
    CheckTrue "Mode de paiement valide (casse ignoree)", IsValidPaymentMode("mobile money"), lngOk, lngTotal, strReport
    CheckTrue "Mode de paiement inconnu refuse", Not IsValidPaymentMode("Bitcoin"), lngOk, lngTotal, strReport
    CheckEq "Devise", CurrencyCode(), "FCFA", lngOk, lngTotal, strReport
    CheckEq "Decimales de la devise", CurrencyDecimals(), 0, lngOk, lngTotal, strReport

    ' ---------- Montants ----------
    SplitTaxIncluded 5000000, curNet, curTax
    CheckTrue "5 000 000 TTC = 4 192 872 HT + 807 128 de taxe", _
        curNet = 4192872 And curTax = 807128, lngOk, lngTotal, strReport
    SplitTaxIncluded 1192500, curNet, curTax
    CheckTrue "1 192 500 TTC = 1 000 000 HT + 192 500 de taxe", _
        curNet = 1000000 And curTax = 192500, lngOk, lngTotal, strReport
    SplitTaxIncluded 750000, curNet, curTax, 0
    CheckTrue "Taux 0 % : HT = TTC, taxe nulle", curNet = 750000 And curTax = 0, lngOk, lngTotal, strReport
    CheckEq "Arrondi de 2,5 : 3 (moitie vers le haut, pas d'arrondi bancaire)", RoundAmount(2.5), 3, lngOk, lngTotal, strReport
    CheckEq "Arrondi de 3,5 : 4", RoundAmount(3.5), 4, lngOk, lngTotal, strReport
    CheckEq "Arrondi de -2,5 : -3", RoundAmount(-2.5), -3, lngOk, lngTotal, strReport

    ' ---------- Numerotation ----------
    CheckEq "Numerotation hors transaction refusee", NumberingError(DOC_INVOICE), ERR_NUMBERING, lngOk, lngTotal, strReport
    CheckEq "Format d'un numero", FormatDocumentNumber(DOC_INVOICE, 2026, 7), "FAC-2026-00007", lngOk, lngTotal, strReport

    lngCounters = CLng(QueryValue("SELECT COUNT(*) FROM [Compteur]"))
    lngPlaces = CLng(QueryValue("SELECT COUNT(*) FROM [Lieu]"))
    BeginTransaction
    blnTransaction = True

    ' Annee 2099 : aucun compteur existant, les numeros commencent a 1
    CheckEq "Premiere facture de l'annee", NextDocumentNumber(DOC_INVOICE, DateSerial(2099, 3, 1)), "FAC-2099-00001", _
        lngOk, lngTotal, strReport
    CheckEq "Facture suivante", NextDocumentNumber(DOC_INVOICE, DateSerial(2099, 12, 31)), "FAC-2099-00002", _
        lngOk, lngTotal, strReport
    CheckEq "Recu : compteur distinct", NextDocumentNumber(DOC_RECEIPT, DateSerial(2099, 5, 1)), "REC-2099-00001", _
        lngOk, lngTotal, strReport
    CheckEq "Nouvelle annee : compteur remis a 1", NextDocumentNumber(DOC_INVOICE, DateSerial(2098, 1, 1)), "FAC-2098-00001", _
        lngOk, lngTotal, strReport
    strNumber = NextDocumentNumber(DOC_REFUND)
    CheckTrue "Annee par defaut : annee en cours", strNumber = "RMB-" & Year(Date) & "-" & Right$(strNumber, 5), _
        lngOk, lngTotal, strReport
    CheckEq "Type de piece inconnu refuse", NumberingError("XYZ"), ERR_NUMBERING, lngOk, lngTotal, strReport

    ' ---------- Marques et modeles ----------
    lngBrand = AddBrand("  TEST Toyota  ")
    CheckEq "Marque creee, nom nettoye", QueryValue("SELECT [Nom] FROM [Marque] WHERE [MarqueID] = ?", lngBrand), _
        "TEST Toyota", lngOk, lngTotal, strReport
    CheckEq "Marque en double refusee (casse ignoree)", RefError("AddBrand", "test toyota"), ERR_REF_DUPLICATE, _
        lngOk, lngTotal, strReport
    CheckEq "Marque sans nom refusee", RefError("AddBrand", "   "), ERR_REF_INVALID, lngOk, lngTotal, strReport
    lngBrand2 = AddBrand("TEST Nissan")
    CheckEq "Renommer vers un nom existant refuse", RefError("RenameBrand", "TEST Toyota", lngBrand2), ERR_REF_DUPLICATE, _
        lngOk, lngTotal, strReport

    lngModel = AddModel(lngBrand, "Corolla")
    CheckEq "Libelle du modele", ModelLabel(lngModel), "TEST Toyota Corolla", lngOk, lngTotal, strReport
    CheckEq "Modele en double pour la marque refuse", RefError("AddModel", "corolla", lngBrand), ERR_REF_DUPLICATE, _
        lngOk, lngTotal, strReport
    CheckTrue "Meme nom de modele accepte pour une autre marque", AddModel(lngBrand2, "Corolla") > 0, lngOk, lngTotal, strReport
    CheckEq "Modele pour une marque inconnue refuse", RefError("AddModel", "X", 999999), ERR_REF_NOT_FOUND, _
        lngOk, lngTotal, strReport
    CheckEq "Marque avec modeles : suppression refusee", RefError("DeleteBrand", "", lngBrand), ERR_REF_IN_USE, _
        lngOk, lngTotal, strReport

    ExecuteInsert "INSERT INTO [Voiture] ([ModeleID], [Chassis], [Statut]) VALUES (?, ?, ?)", _
        lngModel, "TESTREF-CHASSIS-1", CAR_STATUS_IN_STOCK
    CheckEq "Modele utilise par une voiture : suppression refusee", RefError("DeleteModel", "", lngModel), ERR_REF_IN_USE, _
        lngOk, lngTotal, strReport

    ' ---------- Lieux ----------
    lngCity = AddPlace("TEST Douala", PLACE_TYPE_CITY)
    lngDistrict = AddPlace("Akwa", PLACE_TYPE_DISTRICT, lngCity)
    lngStation = AddPlace("Commissariat du 1er", PLACE_TYPE_POLICE_STATION, lngCity)
    CheckTrue "Ville, quartier et commissariat crees", lngCity > 0 And lngDistrict > 0 And lngStation > 0, _
        lngOk, lngTotal, strReport
    CheckEq "Quartier sans ville refuse", RefError("AddPlace", "Bonapriso", 0, PLACE_TYPE_DISTRICT), ERR_REF_INVALID, _
        lngOk, lngTotal, strReport
    CheckEq "Ville avec un parent refusee", RefError("AddPlace", "X", lngCity, PLACE_TYPE_CITY), ERR_REF_INVALID, _
        lngOk, lngTotal, strReport
    CheckEq "Type de lieu inconnu refuse", RefError("AddPlace", "X", 0, "PAYS"), ERR_REF_INVALID, lngOk, lngTotal, strReport
    CheckEq "Quartier en double refuse", RefError("AddPlace", "akwa", lngCity, PLACE_TYPE_DISTRICT), ERR_REF_DUPLICATE, _
        lngOk, lngTotal, strReport
    CheckEq "Ville avec des sous-lieux : suppression refusee", RefError("DeletePlace", "", lngCity), ERR_REF_IN_USE, _
        lngOk, lngTotal, strReport
    DeletePlace lngDistrict
    CheckEq "Quartier supprime", CLng(QueryValue("SELECT COUNT(*) FROM [Lieu] WHERE [LieuID] = ?", lngDistrict)), 0, _
        lngOk, lngTotal, strReport

    ' ---------- Intermediaires ----------
    Set objData = New DataAccess
    Set objInter = New Intermediary
    objInter.Name = "  TEST Demarcheur  "
    objInter.Phone = "+237 699 00 00 00"
    CheckEq "Intermediaire sans ville : 1 erreur", objInter.ValidationErrors().Count, 1, lngOk, lngTotal, strReport
    objInter.CityId = lngStation
    CheckEq "Ville de l'intermediaire qui n'est pas une ville refusee", SaveInterError(objData, objInter), _
        ERR_INTERMEDIARY_INVALID, lngOk, lngTotal, strReport
    objInter.CityId = lngCity
    objData.SaveIntermediary objInter
    CheckTrue "Intermediaire enregistre", objInter.Id > 0, lngOk, lngTotal, strReport
    Set objRead = objData.LoadIntermediary(objInter.Id)
    CheckTrue "Intermediaire relu identique", objRead.Name = "TEST Demarcheur" And objRead.Phone = "+237 699 00 00 00" _
        And objRead.CityId = lngCity, lngOk, lngTotal, strReport
    objInter.Phone = "abc"
    CheckEq "Telephone invalide refuse", SaveInterError(objData, objInter), ERR_INTERMEDIARY_INVALID, lngOk, lngTotal, strReport
    CheckEq "Ville d'un intermediaire : suppression refusee", RefError("DeletePlace", "", lngCity), ERR_REF_IN_USE, _
        lngOk, lngTotal, strReport

    ' ---------- Fin : on annule tout ----------
    RollbackTransaction
    blnTransaction = False

    CheckEq "Compteurs revenus a l'etat initial", CLng(QueryValue("SELECT COUNT(*) FROM [Compteur]")), lngCounters, _
        lngOk, lngTotal, strReport
    CheckEq "Lieux revenus a l'etat initial", CLng(QueryValue("SELECT COUNT(*) FROM [Lieu]")), lngPlaces, _
        lngOk, lngTotal, strReport
    CheckEq "Aucune marque de test restante", _
        CLng(QueryValue("SELECT COUNT(*) FROM [Marque] WHERE [Nom] LIKE ?", "TEST %")), 0, lngOk, lngTotal, strReport

    Debug.Print lngOk & " tests reussis sur " & lngTotal & "." & vbCrLf & strReport   ' rapport complet (Ctrl+G)
    MsgBox lngOk & " tests reussis sur " & lngTotal & "." & vbCrLf & vbCrLf & strReport, _
        IIf(lngOk = lngTotal, vbInformation, vbExclamation), "Kamerix - tests du referentiel"
    Exit Sub

EH:
    lngErr = Err.Number
    strErrText = Err.Description
    On Error Resume Next
    If blnTransaction Then RollbackTransaction
    On Error GoTo 0
    MsgBox "Echec (erreur " & lngErr & ") :" & vbCrLf & vbCrLf & strErrText & vbCrLf & vbCrLf & _
           "Resultats deja obtenus :" & vbCrLf & strReport, vbCritical, "Kamerix - tests du referentiel"
End Sub

' Numero d'erreur de NextDocumentNumber (0 si accepte)
Private Function NumberingError(ByVal pDocType As String) As Long
    On Error Resume Next
    NextDocumentNumber pDocType
    NumberingError = Err.Number
    Err.Clear
End Function

' Numero d'erreur d'une operation du referentiel (0 si acceptee)
Private Function RefError(ByVal pOperation As String, ByVal pName As String, Optional ByVal pId As Long, _
                          Optional ByVal pPlaceType As String) As Long
    On Error Resume Next
    Select Case pOperation
        Case "AddBrand":    AddBrand pName
        Case "RenameBrand": RenameBrand pId, pName
        Case "DeleteBrand": DeleteBrand pId
        Case "AddModel":    AddModel pId, pName
        Case "DeleteModel": DeleteModel pId
        Case "AddPlace":    AddPlace pName, pPlaceType, pId
        Case "DeletePlace": DeletePlace pId
    End Select
    RefError = Err.Number
    Err.Clear
End Function

Private Function SaveInterError(ByVal pData As DataAccess, ByVal pInter As Intermediary) As Long
    On Error Resume Next
    pData.SaveIntermediary pInter
    SaveInterError = Err.Number
    Err.Clear
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
