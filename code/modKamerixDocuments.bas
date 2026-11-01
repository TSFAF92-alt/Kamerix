Attribute VB_Name = "modKamerixDocuments"
Option Explicit
Option Private Module

'==============================================================================
' KAMERIX - Documents imprimables : facture, reçu, reçu de remboursement (lot 5.5)
'------------------------------------------------------------------------------
' Principe    : le document est construit par code dans un classeur temporaire
'               (A4 portrait, mise en page fixée ici), exporté en PDF, éventuellement
'               imprimé, puis le classeur est fermé sans être enregistré.
'               Aucune feuille modèle : rien à protéger ni à maintenir dans l'add-in.
' Documents   : FAC (facture, A4 entier) ; REC (reçu de paiement) et RMB (reçu de
'               remboursement) en A4 avec DEUX exemplaires pliables : moitié haute
'               pour le client, moitié basse pour la souche de l'entreprise.
' Éditions    : la première édition est l'original ; les suivantes portent la mention
'               DUPLICATA n° N. Chaque édition est tracée dans JournalDocument
'               (Action PDF ou IMPRESSION, Duplicata Oui/Non). Une facture annulée
'               reste imprimable avec la mention ANNULÉE en filigrane.
' Fichiers    : <dossier du classeur>\Documents\AAAA\Factures|Recus|Remboursements\
'               FAC-2026-00001.pdf ; duplicata : FAC-2026-00001_DUP2.pdf.
'               Un fichier existant n'est jamais écrasé : " (2)", " (3)"... est ajouté.
' En-tête     : paramètres SOCIETE_* et PIED_PAGE de la table Parametre, créés avec
'               des valeurs d'exemple par EnsureDocumentSettings s'ils manquent
'               (appelé par StartKamerix et avant chaque édition).
' Blocage     : aucune édition tant que l'entreprise n'est pas configurée et validée
'               (CONFIG_ENTREPRISE_FAITE, modKamerixCompany ; lot 5.6).
' Droits      : facture : lecture sur Factures ; reçus : lecture sur Paiements ;
'               liste des documents d'une vente : lecture sur Ventes.
' Données     : DocumentData renvoie un Scripting.Dictionary (testable sans Excel) ;
'               les procédures Build* ne font que la mise en page.
' Dependances : modKamerixSql, modKamerixSession, modKamerixSettings,
'               modKamerixAmountWords, modKamerixConst, modKamerixCompany.
'==============================================================================

Public Const ERR_DOCUMENT As Long = vbObjectError + 1900

' Sortie demandée à ProduceDocument
Public Enum KxDocOutput
    kxDocPdf = 1          ' PDF enregistré puis ouvert
    kxDocPrint = 2        ' PDF enregistré puis imprimé sur l'imprimante par défaut de Windows
    kxDocFileOnly = 3     ' PDF enregistré seulement (tests)
End Enum

' Colonnes renvoyées par DocumentRows (tableau colonne x ligne pour ListBox.Column)
Public Const DOC_COL_ID As Long = 0
Public Const DOC_COL_KIND As Long = 1
Public Const DOC_COL_NUMBER As Long = 2
Public Const DOC_COL_DATE As Long = 3
Public Const DOC_COL_AMOUNT As Long = 4
Public Const DOC_COL_EDITIONS As Long = 5
Public Const DOC_COL_NOTE As Long = 6
Public Const DOC_COL_TYPE As Long = 7
Private Const DOC_COL_COUNT As Long = 8

' Paramètres de l'en-tête des documents (table Parametre)
Public Const PARAM_COMPANY_NAME As String = "SOCIETE_NOM"
Public Const PARAM_COMPANY_ADDRESS As String = "SOCIETE_ADRESSE"
Public Const PARAM_COMPANY_PHONE As String = "SOCIETE_TELEPHONE"
Public Const PARAM_COMPANY_EMAIL As String = "SOCIETE_EMAIL"
Public Const PARAM_COMPANY_NIU As String = "SOCIETE_NIU"
Public Const PARAM_COMPANY_RCCM As String = "SOCIETE_RCCM"
Public Const PARAM_COMPANY_LOGO As String = "SOCIETE_LOGO"
Public Const PARAM_FOOTER As String = "PIED_PAGE"

' Actions tracées dans JournalDocument
Public Const DOC_ACTION_PDF As String = "PDF"
Public Const DOC_ACTION_PRINT As String = "IMPRESSION"

' Mise en page : 6 colonnes égales, A4 portrait (largeur utile 510 pt avec 1,5 cm de marge)
Private Const COL_COUNT As Long = 6
Private Const COL_WIDTH_PT As Double = 84
Private Const FONT_NAME As String = "Calibri"
Private Const FONT_SIZE As Double = 10
Private Const RECEIPT_TOP_BLOCK_PT As Double = 392.25     ' moitié haute : le pli tombe au milieu de la page
Private Const FOLD_ROW_PT As Double = 12
Private Const COLOR_TITLE As Long = 6567712               ' RGB(32, 55, 100) bleu foncé
Private Const COLOR_HEAD_FILL As Long = 6567712
Private Const COLOR_LIGHT_FILL As Long = 16247773         ' RGB(221, 235, 247)
Private Const COLOR_GRAY_FILL As Long = 15921906          ' RGB(242, 242, 242)
Private Const COLOR_GRAY_TEXT As Long = 8421504           ' RGB(128, 128, 128)
Private Const COLOR_RED As Long = 192                     ' RGB(192, 0, 0)

' Dossier racine des PDF imposé par les tests (vide = dossier normal)
Private mstrTestRoot As String

' Cache le temps d'une édition (une seule lecture de la table Parametre)
Private mdicSettings As Object
Private mstrCurrencyCode As String
Private mlngCurrencyDecimals As Long

'==============================================================================
' API PUBLIQUE : EDITION
'==============================================================================

' Produit le document (PDF, et impression si demandée) et retourne le chemin du PDF.
'   pDocType : DOC_INVOICE (pDocId = FactureID), DOC_RECEIPT ou DOC_REFUND (pDocId = PaiementID)
Public Function ProduceDocument(ByVal pDocType As String, ByVal pDocId As Long, _
                                ByVal pOutput As KxDocOutput) As String
    Const SRC As String = "modKamerixDocuments.ProduceDocument"
    Dim dicData As Object
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lngEdition As Long
    Dim strPath As String
    Dim blnScreen As Boolean
    Dim lngErr As Long
    Dim strErr As String
    Dim strSrc As String

    RequireDocumentRead pDocType
    If pOutput <> kxDocPdf And pOutput <> kxDocPrint And pOutput <> kxDocFileOnly Then
        Err.Raise ERR_DOCUMENT, SRC, "Sortie inconnue : " & pOutput & "."
    End If
    EnsureDocumentSettings
    If Not IsCompanyConfigured() Then
        Err.Raise ERR_DOCUMENT, SRC, "Les documents ne peuvent pas " & ChrW(234) & "tre " & ChrW(233) & "dit" & ChrW(233) & _
            "s tant que l'entreprise n'est pas configur" & ChrW(233) & "e." & vbCrLf & _
            "Un administrateur doit la renseigner : OpenScreen ""PARAMETRES"", onglet Entreprise."
    End If

    Set dicData = DocumentData(pDocType, pDocId)
    blnScreen = Application.ScreenUpdating
    BeginSettingsCache
    On Error GoTo EH                                ' le cache est toujours libéré, même en cas d'erreur

    lngEdition = EditionCount(pDocType, pDocId) + 1
    strPath = UniqueFilePath(DocumentFolder(pDocType, dicData("Date")), PdfBaseName(dicData("Number"), lngEdition))

    Application.ScreenUpdating = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    If pDocType = DOC_INVOICE Then
        BuildInvoiceSheet ws, dicData, lngEdition
    Else
        BuildReceiptSheet ws, dicData, lngEdition
    End If

    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=strPath, Quality:=xlQualityStandard, _
                           IncludeDocProperties:=True, IgnorePrintAreas:=False, _
                           OpenAfterPublish:=(pOutput = kxDocPdf)
    If pOutput = kxDocPrint Then ws.PrintOut Copies:=1

    wb.Close SaveChanges:=False
    Set wb = Nothing
    Application.ScreenUpdating = blnScreen
    EndSettingsCache

    LogEdition pDocType, pDocId, IIf(pOutput = kxDocPrint, DOC_ACTION_PRINT, DOC_ACTION_PDF), (lngEdition > 1)
    ProduceDocument = strPath
    Exit Function

EH:
    lngErr = Err.Number: strErr = Err.Description: strSrc = Err.Source
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Application.ScreenUpdating = blnScreen
    EndSettingsCache
    On Error GoTo 0
    Err.Raise lngErr, strSrc, strErr
End Function

' Vrai si le rôle connecté peut éditer ce type de document (pour griser un bouton)
Public Function CanReadDocument(ByVal pDocType As String) As Boolean
    If pDocType = DOC_INVOICE Then
        CanReadDocument = CanRead(MOD_INVOICES)
    ElseIf pDocType = DOC_RECEIPT Or pDocType = DOC_REFUND Then
        CanReadDocument = CanRead(MOD_PAYMENTS)
    End If
End Function

' Nombre d'éditions déjà faites d'un document (0 : la prochaine sera l'original)
Public Function EditionCount(ByVal pDocType As String, ByVal pDocId As Long) As Long
    EditionCount = CLng(QueryValue("SELECT COUNT(*) FROM [JournalDocument] WHERE [TypeDocument] = ? AND [DocumentID] = ?", _
                                   pDocType, pDocId))
End Function

' Identifiant du paiement qui porte ce numéro de pièce (REC-... ou RMB-...) ; 0 si inconnu
Public Function PaymentIdOfReceipt(ByVal pReceiptNumber As String) As Long
    Dim vId As Variant

    vId = QueryValue("SELECT [PaiementID] FROM [Paiement] WHERE [NumeroRecu] = ?", Trim$(pReceiptNumber))
    If Not IsNull(vId) Then PaymentIdOfReceipt = CLng(vId)
End Function

' Libellé d'un type de document
Public Function DocumentKindLabel(ByVal pDocType As String) As String
    Select Case pDocType
        Case DOC_INVOICE: DocumentKindLabel = "Facture"
        Case DOC_RECEIPT: DocumentKindLabel = "Reçu"
        Case DOC_REFUND:  DocumentKindLabel = "Remboursement"
        Case Else:        DocumentKindLabel = pDocType
    End Select
End Function

'==============================================================================
' API PUBLIQUE : LISTE DES DOCUMENTS D'UNE VENTE (écran frmSaleDocuments)
'==============================================================================

' Facture puis paiements de la vente (ordre chronologique), prêts pour ListBox.Column.
' Empty si la vente n'a encore aucun document. Lecture : module Ventes.
Public Function DocumentRows(ByVal pMovementId As Long) As Variant
    Dim rsInvoices As ADODB.Recordset
    Dim rsPayments As ADODB.Recordset
    Dim vData() As Variant
    Dim lngRows As Long
    Dim i As Long
    Dim strType As String

    RequireRead MOD_SALES

    Set rsInvoices = QueryRecordset("SELECT [FactureID], [NumeroFacture], [DateFacture], [MontantTTC], [Statut] " & _
                                    "FROM [Facture] WHERE [MouvementID] = ? ORDER BY [FactureID]", pMovementId)
    Set rsPayments = QueryRecordset("SELECT [PaiementID], [NumeroRecu], [DatePaiement], [Montant], [Sens] " & _
                                    "FROM [Paiement] WHERE [MouvementID] = ? ORDER BY [DatePaiement], [PaiementID]", pMovementId)
    lngRows = RecordCount(rsInvoices) + RecordCount(rsPayments)
    If lngRows = 0 Then
        rsInvoices.Close
        rsPayments.Close
        Exit Function
    End If

    ReDim vData(0 To DOC_COL_COUNT - 1, 0 To lngRows - 1)
    Do While Not rsInvoices.EOF
        vData(DOC_COL_ID, i) = CLng(rsInvoices.Fields("FactureID").Value)
        vData(DOC_COL_TYPE, i) = DOC_INVOICE
        vData(DOC_COL_KIND, i) = DocumentKindLabel(DOC_INVOICE)
        vData(DOC_COL_NUMBER, i) = CStr(rsInvoices.Fields("NumeroFacture").Value)
        vData(DOC_COL_DATE, i) = Format$(rsInvoices.Fields("DateFacture").Value, "dd/mm/yyyy")
        vData(DOC_COL_AMOUNT, i) = Format$(rsInvoices.Fields("MontantTTC").Value, "#,##0")
        vData(DOC_COL_EDITIONS, i) = CStr(EditionCount(DOC_INVOICE, CLng(vData(DOC_COL_ID, i))))
        If CStr(rsInvoices.Fields("Statut").Value) = INVOICE_STATUS_CANCELLED Then vData(DOC_COL_NOTE, i) = "Annulée" Else vData(DOC_COL_NOTE, i) = ""
        i = i + 1
        rsInvoices.MoveNext
    Loop
    rsInvoices.Close

    Do While Not rsPayments.EOF
        If CStr(rsPayments.Fields("Sens").Value) = PAYMENT_DIRECTION_REFUND Then strType = DOC_REFUND Else strType = DOC_RECEIPT
        vData(DOC_COL_ID, i) = CLng(rsPayments.Fields("PaiementID").Value)
        vData(DOC_COL_TYPE, i) = strType
        vData(DOC_COL_KIND, i) = DocumentKindLabel(strType)
        vData(DOC_COL_NUMBER, i) = CStr(rsPayments.Fields("NumeroRecu").Value)
        vData(DOC_COL_DATE, i) = Format$(rsPayments.Fields("DatePaiement").Value, "dd/mm/yyyy")
        vData(DOC_COL_AMOUNT, i) = Format$(rsPayments.Fields("Montant").Value, "#,##0")
        vData(DOC_COL_EDITIONS, i) = CStr(EditionCount(strType, CLng(vData(DOC_COL_ID, i))))
        vData(DOC_COL_NOTE, i) = ""
        i = i + 1
        rsPayments.MoveNext
    Loop
    rsPayments.Close

    DocumentRows = vData
End Function

'==============================================================================
' API PUBLIQUE : PARAMETRES DE L'EN-TETE
'==============================================================================

' Crée les paramètres SOCIETE_* et PIED_PAGE absents, avec des valeurs d'exemple à
' remplacer par les vraies (table Parametre). Ne modifie jamais un paramètre existant.
Public Sub EnsureDocumentSettings()
    Dim dicExisting As Object
    Dim vDefaults As Variant
    Dim i As Long

    Set dicExisting = SettingsMap()
    vDefaults = DocumentSettingDefaults()
    For i = LBound(vDefaults) To UBound(vDefaults) Step 2
        If Not dicExisting.Exists(UCase$(vDefaults(i))) Then
            ExecuteNonQuery "INSERT INTO [Parametre] ([Cle], [Valeur]) VALUES (?, ?)", vDefaults(i), NullIfEmpty(vDefaults(i + 1))
        End If
    Next i
End Sub

' Valeur d'un paramètre de l'en-tête ("" s'il est absent ou vide)
Public Function CompanySetting(ByVal pKey As String) As String
    Dim vValue As Variant

    If Not mdicSettings Is Nothing Then
        If mdicSettings.Exists(UCase$(pKey)) Then CompanySetting = mdicSettings(UCase$(pKey))
        Exit Function
    End If
    vValue = QueryValue("SELECT [Valeur] FROM [Parametre] WHERE [Cle] = ?", pKey)
    If Not IsNull(vValue) Then CompanySetting = Trim$(CStr(vValue))
End Function

' Paires clé, valeur d'exemple (à modifier dans la table Parametre)
Public Function DocumentSettingDefaults() As Variant
    DocumentSettingDefaults = Array( _
        PARAM_COMPANY_NAME, "MA SOCIETE AUTO SARL", _
        PARAM_COMPANY_ADDRESS, "Rue de l'Exemple, Akwa - BP 0000 Douala", _
        PARAM_COMPANY_PHONE, "+237 6 00 00 00 00", _
        PARAM_COMPANY_EMAIL, "contact@exemple.cm", _
        PARAM_COMPANY_NIU, "M000000000000X", _
        PARAM_COMPANY_RCCM, "RC/DLA/2026/B/0000", _
        PARAM_COMPANY_LOGO, "", _
        PARAM_FOOTER, "Merci de votre confiance.")
End Function

'==============================================================================
' API PUBLIQUE : DONNEES D'UN DOCUMENT (sans mise en page, testables)
'==============================================================================

' Données d'un document dans un Scripting.Dictionary. Clés communes :
'   Type, Id, Number, Date, Status, MovementId, Price, PartyName, PartyPhone, PartyAddress,
'   PartyIdDoc, CarLabel, Chassis, Plate, Color, Mileage, Issuer, AmountWords
' Facture : Designation, Net, TaxRate, Tax, Total, PaymentRows (tableau 4 x n ou Empty), PaidTotal
' Reçu    : Amount, Mode, PaidToDate, Balance, InvoiceNumber ("" sauf sur le reçu qui solde)
' Remboursement : Amount, Mode, PaidBefore, Withheld, Reason
Public Function DocumentData(ByVal pDocType As String, ByVal pDocId As Long) As Object
    RequireDocumentRead pDocType
    Select Case pDocType
        Case DOC_INVOICE
            Set DocumentData = InvoiceData(pDocId)
        Case DOC_RECEIPT, DOC_REFUND
            Set DocumentData = PaymentData(pDocType, pDocId)
    End Select
End Function

'==============================================================================
' API PUBLIQUE : OUTILS DE TEST
'==============================================================================

' Dossier racine des PDF pendant les tests ("" pour revenir au dossier normal)
Public Sub SetDocumentFolderForTests(ByVal pFolder As String)
    mstrTestRoot = pFolder
End Sub

'==============================================================================
' DONNEES (PROCEDURES INTERNES)
'==============================================================================

Private Function InvoiceData(ByVal pInvoiceId As Long) As Object
    Const SRC As String = "modKamerixDocuments.InvoiceData"
    Dim d As Object
    Dim rs As ADODB.Recordset
    Dim vRows As Variant
    Dim curPaid As Currency
    Dim i As Long

    Set rs = QueryRecordset( _
        "SELECT f.[NumeroFacture], f.[DateFacture], f.[Designation], f.[MontantHT], f.[TauxTaxe], f.[MontantTaxe], " & _
        "f.[MontantTTC], f.[Statut], f.[MouvementID], t.[Nom], t.[Telephone], t.[Adresse], t.[TypePiece], t.[NumeroPiece], " & _
        "u.[NomComplet] " & _
        "FROM ([Facture] AS f INNER JOIN [Tiers] AS t ON f.[TiersID] = t.[TiersID]) " & _
        "INNER JOIN [Utilisateur] AS u ON f.[UtilisateurID] = u.[UtilisateurID] " & _
        "WHERE f.[FactureID] = ?", pInvoiceId)
    If rs.EOF Then
        rs.Close
        Err.Raise ERR_DOCUMENT, SRC, "Facture introuvable (identifiant " & pInvoiceId & ")."
    End If

    Set d = CreateObject("Scripting.Dictionary")
    d("Type") = DOC_INVOICE
    d("Id") = pInvoiceId
    d("Number") = CStr(rs.Fields("NumeroFacture").Value)
    d("Date") = CDate(rs.Fields("DateFacture").Value)
    d("Status") = CStr(rs.Fields("Statut").Value)
    d("MovementId") = CLng(rs.Fields("MouvementID").Value)
    d("Designation") = CStr(rs.Fields("Designation").Value)
    d("Net") = CCur(rs.Fields("MontantHT").Value)
    d("TaxRate") = CDbl(rs.Fields("TauxTaxe").Value)
    d("Tax") = CCur(rs.Fields("MontantTaxe").Value)
    d("Total") = CCur(rs.Fields("MontantTTC").Value)
    d("Issuer") = CStr(rs.Fields("NomComplet").Value)
    AddPartyFields d, rs
    rs.Close

    AddCarFields d, d("MovementId")
    d("AmountWords") = AmountInWords(d("Total"))

    ' Règlements rattachés à la facture (encaissements)
    Set rs = QueryRecordset("SELECT [NumeroRecu], [DatePaiement], [ModePaiement], [Montant] FROM [Paiement] " & _
                            "WHERE [FactureID] = ? AND [Sens] = ? ORDER BY [DatePaiement], [PaiementID]", _
                            pInvoiceId, PAYMENT_DIRECTION_IN)
    If rs.EOF Then
        rs.Close
        d("PaymentRows") = Empty
    Else
        vRows = rs.GetRows()
        rs.Close
        For i = 0 To UBound(vRows, 2)
            curPaid = curPaid + CCur(vRows(3, i))
        Next i
        d("PaymentRows") = vRows
    End If
    d("PaidTotal") = curPaid

    Set InvoiceData = d
End Function

Private Function PaymentData(ByVal pDocType As String, ByVal pPaymentId As Long) As Object
    Const SRC As String = "modKamerixDocuments.PaymentData"
    Dim d As Object
    Dim rs As ADODB.Recordset
    Dim strDirection As String
    Dim curIn As Currency
    Dim curOut As Currency
    Dim vInvoice As Variant

    Set rs = QueryRecordset( _
        "SELECT p.[NumeroRecu], p.[Sens], p.[DatePaiement], p.[Montant], p.[ModePaiement], p.[MouvementID], p.[FactureID], " & _
        "t.[Nom], t.[Telephone], t.[Adresse], t.[TypePiece], t.[NumeroPiece], u.[NomComplet] " & _
        "FROM (([Paiement] AS p INNER JOIN [Mouvement] AS mv ON p.[MouvementID] = mv.[MouvementID]) " & _
        "INNER JOIN [Tiers] AS t ON mv.[TiersID] = t.[TiersID]) " & _
        "INNER JOIN [Utilisateur] AS u ON p.[UtilisateurID] = u.[UtilisateurID] " & _
        "WHERE p.[PaiementID] = ?", pPaymentId)
    If rs.EOF Then
        rs.Close
        Err.Raise ERR_DOCUMENT, SRC, "Pi" & ChrW(232) & "ce introuvable ou non rattach" & ChrW(233) & "e " & ChrW(224) & _
            " une vente (paiement " & pPaymentId & ")."
    End If

    strDirection = CStr(rs.Fields("Sens").Value)
    If (pDocType = DOC_RECEIPT And strDirection <> PAYMENT_DIRECTION_IN) Or _
       (pDocType = DOC_REFUND And strDirection <> PAYMENT_DIRECTION_REFUND) Then
        rs.Close
        Err.Raise ERR_DOCUMENT, SRC, "Le paiement " & pPaymentId & " n'est pas un " & _
            IIf(pDocType = DOC_RECEIPT, "encaissement", "remboursement") & "."
    End If

    Set d = CreateObject("Scripting.Dictionary")
    d("Type") = pDocType
    d("Id") = pPaymentId
    d("Number") = CStr(rs.Fields("NumeroRecu").Value)
    d("Date") = CDate(rs.Fields("DatePaiement").Value)
    d("Status") = strDirection
    d("Amount") = CCur(rs.Fields("Montant").Value)
    d("Mode") = CStr(rs.Fields("ModePaiement").Value)
    d("MovementId") = CLng(rs.Fields("MouvementID").Value)
    vInvoice = rs.Fields("FactureID").Value
    d("Issuer") = CStr(rs.Fields("NomComplet").Value)
    AddPartyFields d, rs
    rs.Close

    AddCarFields d, d("MovementId")
    d("AmountWords") = AmountInWords(d("Amount"))

    If pDocType = DOC_RECEIPT Then
        ' Situation de la vente juste après ce paiement (pièces enregistrées jusqu'à lui)
        curIn = SumPayments(d("MovementId"), PAYMENT_DIRECTION_IN, pPaymentId, True)
        curOut = SumPayments(d("MovementId"), PAYMENT_DIRECTION_REFUND, pPaymentId, True)
        d("PaidToDate") = curIn - curOut
        d("Balance") = d("Price") - d("PaidToDate")
        If d("Balance") < 0 Then d("Balance") = CCur(0)
        d("InvoiceNumber") = ""
        If d("Balance") = 0 And Not IsNull(vInvoice) Then
            d("InvoiceNumber") = TextOf(QueryValue("SELECT [NumeroFacture] FROM [Facture] WHERE [FactureID] = ?", CLng(vInvoice)))
        End If
    Else
        ' Sommes versées avant le remboursement ; retenue = versé - remboursé
        curIn = SumPayments(d("MovementId"), PAYMENT_DIRECTION_IN, pPaymentId, False)
        curOut = SumPayments(d("MovementId"), PAYMENT_DIRECTION_REFUND, pPaymentId, False)
        d("PaidBefore") = curIn - curOut
        d("Withheld") = d("PaidBefore") - d("Amount")
        If d("Withheld") < 0 Then d("Withheld") = CCur(0)
    End If

    Set PaymentData = d
End Function

' Total des paiements d'un sens jusqu'au paiement pPaymentId (inclus ou non)
Private Function SumPayments(ByVal pMovementId As Long, ByVal pDirection As String, ByVal pPaymentId As Long, _
                             ByVal pIncluded As Boolean) As Currency
    Dim vSum As Variant

    vSum = QueryValue("SELECT SUM([Montant]) FROM [Paiement] WHERE [MouvementID] = ? AND [Sens] = ? AND [PaiementID] " & _
                      IIf(pIncluded, "<=", "<") & " ?", pMovementId, pDirection, pPaymentId)
    If Not IsNull(vSum) Then SumPayments = CCur(vSum)
End Function

' Client : nom, téléphone, adresse, pièce d'identité (champs Nom, Telephone, Adresse, TypePiece, NumeroPiece)
Private Sub AddPartyFields(ByVal d As Object, ByVal rs As ADODB.Recordset)
    Dim strDoc As String

    d("PartyName") = TextOf(rs.Fields("Nom").Value)
    d("PartyPhone") = TextOf(rs.Fields("Telephone").Value)
    d("PartyAddress") = TextOf(rs.Fields("Adresse").Value)
    strDoc = TextOf(rs.Fields("TypePiece").Value)
    If Len(TextOf(rs.Fields("NumeroPiece").Value)) > 0 Then
        If Len(strDoc) = 0 Then strDoc = "Pi" & ChrW(232) & "ce"
        strDoc = strDoc & " n" & ChrW(176) & " " & TextOf(rs.Fields("NumeroPiece").Value)
    End If
    d("PartyIdDoc") = strDoc
End Sub

' Voiture et prix de la vente, motif d'annulation
Private Sub AddCarFields(ByVal d As Object, ByVal pMovementId As Long)
    Dim rs As ADODB.Recordset
    Dim strLabel As String

    Set rs = QueryRecordset( _
        "SELECT mv.[Prix], mv.[MotifAnnulation], b.[Nom] AS Marque, m.[Nom] AS Modele, v.[Annee], v.[Couleur], " & _
        "v.[Chassis], v.[Immatriculation], v.[Kilometrage] " & _
        "FROM (([Mouvement] AS mv INNER JOIN [Voiture] AS v ON mv.[VoitureID] = v.[VoitureID]) " & _
        "INNER JOIN [Modele] AS m ON v.[ModeleID] = m.[ModeleID]) " & _
        "INNER JOIN [Marque] AS b ON m.[MarqueID] = b.[MarqueID] " & _
        "WHERE mv.[MouvementID] = ?", pMovementId)
    If rs.EOF Then
        rs.Close
        Err.Raise ERR_DOCUMENT, "modKamerixDocuments.AddCarFields", "Vente introuvable (identifiant " & pMovementId & ")."
    End If

    strLabel = TextOf(rs.Fields("Marque").Value) & " " & TextOf(rs.Fields("Modele").Value)
    If Not IsNull(rs.Fields("Annee").Value) Then strLabel = strLabel & " " & CStr(rs.Fields("Annee").Value)
    d("Price") = CCur(rs.Fields("Prix").Value)
    d("Reason") = TextOf(rs.Fields("MotifAnnulation").Value)
    d("CarLabel") = strLabel
    d("Chassis") = TextOf(rs.Fields("Chassis").Value)
    d("Plate") = TextOf(rs.Fields("Immatriculation").Value)
    d("Color") = TextOf(rs.Fields("Couleur").Value)
    If IsNull(rs.Fields("Kilometrage").Value) Then d("Mileage") = "" Else d("Mileage") = Format$(rs.Fields("Kilometrage").Value, "#,##0") & " km"
    rs.Close
End Sub

'==============================================================================
' MISE EN PAGE : FACTURE
'==============================================================================

Private Sub BuildInvoiceSheet(ByVal ws As Worksheet, ByVal d As Object, ByVal pEdition As Long)
    Dim r As Long
    Dim rTable As Long
    Dim rTotal As Long
    Dim vRows As Variant
    Dim i As Long
    Dim strLine As String

    PrepareSheet ws

    ' ---------- En-tête : entreprise à gauche, titre à droite ----------
    r = WriteCompanyHeader(ws, 1, 4, True)
    PutText(ws, 1, 5, 6, "FACTURE", True, 20, xlRight).Font.Color = COLOR_TITLE
    PutText ws, 2, 5, 6, "N" & ChrW(176) & " " & d("Number"), True, 11, xlRight
    PutText ws, 3, 5, 6, "Date : " & Format$(d("Date"), "dd/mm/yyyy"), False, 10, xlRight
    If pEdition > 1 Then PutText(ws, 4, 5, 6, DuplicateText(pEdition), True, 10, xlRight).Font.Color = COLOR_RED
    r = 5
    SetRowHeight ws, r, 8
    ws.Range(ws.Cells(r, 1), ws.Cells(r, COL_COUNT)).Borders(xlEdgeBottom).Weight = xlMedium
    ws.Range(ws.Cells(r, 1), ws.Cells(r, COL_COUNT)).Borders(xlEdgeBottom).Color = COLOR_TITLE

    ' ---------- Objet (gauche) et client (droite, encadré) ----------
    r = 7
    PutText(ws, r, 1, 3, "OBJET", True, 8, xlLeft).Font.Color = COLOR_GRAY_TEXT
    PutText(ws, r, 4, 6, "FACTUR" & ChrW(201) & " " & ChrW(192), True, 8, xlLeft).Font.Color = COLOR_GRAY_TEXT
    PutText ws, r + 1, 1, 3, "Vente d'un v" & ChrW(233) & "hicule d'occasion", False, 10, xlLeft
    PutText ws, r + 2, 1, 3, "R" & ChrW(232) & "glement : comptant", False, 10, xlLeft
    PutText ws, r + 1, 4, 6, d("PartyName"), True, 11, xlLeft
    PutText ws, r + 2, 4, 6, d("PartyAddress"), False, 10, xlLeft
    PutText ws, r + 3, 4, 6, PrefixIfAny("T" & ChrW(233) & "l. : ", d("PartyPhone")), False, 10, xlLeft
    PutText ws, r + 4, 4, 6, d("PartyIdDoc"), False, 10, xlLeft
    BoxRange ws.Range(ws.Cells(r, 4), ws.Cells(r + 4, 6))

    ' ---------- Tableau de la désignation ----------
    r = r + 6
    rTable = r
    PutText ws, r, 1, 4, "D" & ChrW(233) & "signation", True, 10, xlLeft
    PutText ws, r, 5, 5, "Quantit" & ChrW(233), True, 10, xlCenter
    PutText ws, r, 6, 6, "Montant HT", True, 10, xlRight
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, COL_COUNT))
        .Interior.Color = COLOR_HEAD_FILL
        .Font.Color = vbWhite
    End With
    SetRowHeight ws, r, 18

    r = r + 1
    PutText(ws, r, 1, 4, d("Designation"), False, 10, xlLeft).WrapText = True
    PutText ws, r, 5, 5, "1", False, 10, xlCenter
    PutAmount ws, r, 6, 6, d("Net"), False
    SetRowHeight ws, r, 32
    ws.Range(ws.Cells(r, 1), ws.Cells(r, COL_COUNT)).VerticalAlignment = xlTop
    If Len(d("Mileage")) > 0 Then
        r = r + 1
        PutText(ws, r, 1, 4, "Kilom" & ChrW(233) & "trage : " & d("Mileage"), False, 9, xlLeft).Font.Color = COLOR_GRAY_TEXT
    End If
    BoxRange ws.Range(ws.Cells(rTable, 1), ws.Cells(r, COL_COUNT))

    ' ---------- Totaux ----------
    r = r + 2
    rTotal = r
    PutText ws, r, 4, 5, "Total HT", False, 10, xlRight
    PutAmount ws, r, 6, 6, d("Net"), False
    PutText ws, r + 1, 4, 5, "TVA (" & RateText(d("TaxRate")) & " %)", False, 10, xlRight
    PutAmount ws, r + 1, 6, 6, d("Tax"), False
    PutText ws, r + 2, 4, 5, "Total TTC", True, 11, xlRight
    PutAmount ws, r + 2, 6, 6, d("Total"), True
    ws.Range(ws.Cells(r + 2, 4), ws.Cells(r + 2, 6)).Interior.Color = COLOR_LIGHT_FILL
    SetRowHeight ws, r + 2, 18
    BoxRange ws.Range(ws.Cells(rTotal, 4), ws.Cells(r + 2, 6))

    ' ---------- Montant en lettres ----------
    r = r + 4
    strLine = "Arr" & ChrW(234) & "t" & ChrW(233) & "e la pr" & ChrW(233) & "sente facture " & ChrW(224) & " la somme de " & _
              d("AmountWords") & " toutes taxes comprises."
    With PutText(ws, r, 1, 6, strLine, False, 10, xlLeft)
        .WrapText = True
        .Font.Italic = True
        .VerticalAlignment = xlTop
    End With
    SetRowHeight ws, r, 30

    ' ---------- Règlements reçus ----------
    r = r + 2
    PutText ws, r, 1, 6, "R" & ChrW(232) & "glements re" & ChrW(231) & "us", True, 10, xlLeft
    r = r + 1
    PutText ws, r, 1, 2, "N" & ChrW(176) & " de re" & ChrW(231) & "u", True, 9, xlLeft
    PutText ws, r, 3, 3, "Date", True, 9, xlCenter
    PutText ws, r, 4, 5, "Mode", True, 9, xlLeft
    PutText ws, r, 6, 6, "Montant", True, 9, xlRight
    ws.Range(ws.Cells(r, 1), ws.Cells(r, COL_COUNT)).Interior.Color = COLOR_GRAY_FILL
    rTable = r
    vRows = d("PaymentRows")
    If Not IsEmpty(vRows) Then
        For i = 0 To UBound(vRows, 2)
            r = r + 1
            PutText ws, r, 1, 2, CStr(vRows(0, i)), False, 9, xlLeft
            PutText ws, r, 3, 3, Format$(vRows(1, i), "dd/mm/yyyy"), False, 9, xlCenter
            PutText ws, r, 4, 5, CStr(vRows(2, i)), False, 9, xlLeft
            PutAmount ws, r, 6, 6, CCur(vRows(3, i)), False
        Next i
    End If
    r = r + 1
    PutText ws, r, 4, 5, "Total r" & ChrW(233) & "gl" & ChrW(233), True, 9, xlRight
    PutAmount ws, r, 6, 6, d("PaidTotal"), True
    BoxRange ws.Range(ws.Cells(rTable, 1), ws.Cells(r, COL_COUNT))

    ' ---------- Signature ----------
    r = r + 3
    PutText ws, r, 4, 6, "Pour " & CompanySetting(PARAM_COMPANY_NAME), True, 10, xlCenter
    PutText ws, r + 1, 4, 6, "Le vendeur : " & d("Issuer"), False, 9, xlCenter
    SetRowHeight ws, r + 2, 48
    r = r + 2

    ' ---------- Filigrane, pied de page, impression ----------
    If d("Status") = INVOICE_STATUS_CANCELLED Then
        AddWatermark ws, "ANNUL" & ChrW(201) & "E", ws.Rows(rTotal - 6).Top, 160
    End If
    ApplyPageSetup ws, r, 1.5, 2, FooterText()
End Sub

'==============================================================================
' MISE EN PAGE : RECU ET REMBOURSEMENT (2 exemplaires pliables sur un A4)
'==============================================================================

Private Sub BuildReceiptSheet(ByVal ws As Worksheet, ByVal d As Object, ByVal pEdition As Long)
    Dim r As Long
    Dim dblUsed As Double

    PrepareSheet ws

    ' Moitié haute : exemplaire du client, complétée jusqu'au pli
    r = WriteReceiptCopy(ws, 1, d, pEdition, "Exemplaire client")
    dblUsed = ws.Range(ws.Rows(1), ws.Rows(r - 1)).Height
    If dblUsed < RECEIPT_TOP_BLOCK_PT Then
        SetRowHeight ws, r, RECEIPT_TOP_BLOCK_PT - dblUsed
        r = r + 1
    End If

    ' Ligne de pliage, puis moitié basse : souche de l'entreprise
    With PutText(ws, r, 1, COL_COUNT, "Plier ou d" & ChrW(233) & "couper suivant le pointill" & ChrW(233), False, 7, xlCenter)
        .Font.Color = COLOR_GRAY_TEXT
        .VerticalAlignment = xlBottom
    End With
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, COL_COUNT)).Borders(xlEdgeTop)
        .LineStyle = xlDash
        .Weight = xlThin
        .Color = COLOR_GRAY_TEXT
    End With
    SetRowHeight ws, r, FOLD_ROW_PT
    r = WriteReceiptCopy(ws, r + 1, d, pEdition, "Souche - exemplaire de l'entreprise")

    ApplyPageSetup ws, r - 1, 1, 1, ""
End Sub

' Un exemplaire du reçu à partir de la ligne pRow ; retourne la première ligne libre après lui
Private Function WriteReceiptCopy(ByVal ws As Worksheet, ByVal pRow As Long, ByVal d As Object, _
                                  ByVal pEdition As Long, ByVal pCopyLabel As String) As Long
    Dim r As Long
    Dim blnRefund As Boolean
    Dim strTitle As String
    Dim strObject As String

    blnRefund = (d("Type") = DOC_REFUND)
    If blnRefund Then
        strTitle = "RE" & ChrW(199) & "U DE REMBOURSEMENT"
    Else
        strTitle = "RE" & ChrW(199) & "U DE PAIEMENT"
    End If

    ' ---------- En-tête ----------
    r = pRow
    WriteCompanyHeader ws, r, 3, False
    PutText(ws, r, 4, 6, strTitle, True, 13, xlRight).Font.Color = COLOR_TITLE
    PutText ws, r + 1, 4, 6, "N" & ChrW(176) & " " & d("Number"), True, 10, xlRight
    PutText ws, r + 2, 4, 6, "Date : " & Format$(d("Date"), "dd/mm/yyyy"), False, 9, xlRight
    With PutText(ws, r + 3, 4, 6, pCopyLabel, False, 8, xlRight)
        .Font.Italic = True
        .Font.Color = COLOR_GRAY_TEXT
    End With
    r = r + 4
    If pEdition > 1 Then PutText(ws, r, 1, 6, DuplicateText(pEdition), True, 9, xlCenter).Font.Color = COLOR_RED
    SetRowHeight ws, r, 13
    r = r + 1
    SetRowHeight ws, r, 4
    ws.Range(ws.Cells(r, 1), ws.Cells(r, COL_COUNT)).Borders(xlEdgeBottom).Weight = xlMedium
    ws.Range(ws.Cells(r, 1), ws.Cells(r, COL_COUNT)).Borders(xlEdgeBottom).Color = COLOR_TITLE
    r = r + 1
    SetRowHeight ws, r, 6

    ' ---------- Corps ----------
    r = r + 1
    PutLabelValue ws, r, IIf(blnRefund, "Rembours" & ChrW(233) & " " & ChrW(224) & " :", "Re" & ChrW(231) & "u de :"), d("PartyName"), True
    r = r + 1
    PutLabelValue ws, r, "T" & ChrW(233) & "l" & ChrW(233) & "phone :", d("PartyPhone"), False
    r = r + 1
    PutText ws, r, 1, 2, "La somme de :", False, 10, xlLeft
    With PutText(ws, r, 3, 6, FormatAmountText(d("Amount")), True, 13, xlLeft)
        .Font.Color = COLOR_TITLE
    End With
    SetRowHeight ws, r, 20
    r = r + 1
    PutText ws, r, 1, 2, "En lettres :", False, 10, xlLeft
    With PutText(ws, r, 3, 6, CapitalizeFirst(d("AmountWords")), False, 10, xlLeft)
        .WrapText = True
        .Font.Italic = True
        .VerticalAlignment = xlTop
    End With
    ws.Cells(r, 1).VerticalAlignment = xlTop
    SetRowHeight ws, r, 28
    r = r + 1
    PutLabelValue ws, r, "Mode :", d("Mode"), False

    strObject = d("CarLabel") & " - ch" & ChrW(226) & "ssis " & d("Chassis")
    If Len(d("Plate")) > 0 Then strObject = strObject & " - " & d("Plate")
    If blnRefund Then
        strObject = "Annulation de la vente : " & strObject
    Else
        strObject = "Paiement sur la vente : " & strObject
    End If
    r = r + 1
    PutText ws, r, 1, 2, "Objet :", False, 10, xlLeft
    With PutText(ws, r, 3, 6, strObject, False, 10, xlLeft)
        .WrapText = True
        .VerticalAlignment = xlTop
    End With
    ws.Cells(r, 1).VerticalAlignment = xlTop
    SetRowHeight ws, r, 26
    If blnRefund And Len(d("Reason")) > 0 Then
        r = r + 1
        PutLabelValue ws, r, "Motif :", d("Reason"), False
    End If

    ' ---------- Situation (3 cases) ----------
    r = r + 1
    SetRowHeight ws, r, 6
    r = r + 1
    If blnRefund Then
        PutSummaryBox ws, r, 1, "Sommes vers" & ChrW(233) & "es", d("PaidBefore")
        PutSummaryBox ws, r, 3, "Retenue (frais du CV...)", d("Withheld")
        PutSummaryBox ws, r, 5, "Montant rembours" & ChrW(233), d("Amount")
    Else
        PutSummaryBox ws, r, 1, "Prix de vente", d("Price")
        PutSummaryBox ws, r, 3, "Total encaiss" & ChrW(233) & " (ce re" & ChrW(231) & "u inclus)", d("PaidToDate")
        PutSummaryBox ws, r, 5, "Reste " & ChrW(224) & " payer", d("Balance")
    End If
    SetRowHeight ws, r, 24
    SetRowHeight ws, r + 1, 18
    BoxRange ws.Range(ws.Cells(r, 1), ws.Cells(r + 1, COL_COUNT))
    r = r + 2
    If Not blnRefund Then
        If Len(d("InvoiceNumber")) > 0 Then
            PutText(ws, r, 1, 6, "Vente sold" & ChrW(233) & "e : facture " & d("InvoiceNumber") & " " & ChrW(233) & "mise.", _
                    False, 9, xlLeft).Font.Italic = True
        End If
    End If
    SetRowHeight ws, r, 14

    ' ---------- Signatures ----------
    r = r + 1
    SetRowHeight ws, r, 6
    r = r + 1
    PutText ws, r, 1, 3, IIf(blnRefund, "Le client (bon pour r" & ChrW(233) & "ception)", "Le client"), True, 9, xlCenter
    PutText ws, r, 4, 6, "Le caissier", True, 9, xlCenter
    r = r + 1
    SetRowHeight ws, r, 34
    r = r + 1
    PutText(ws, r, 4, 6, ChrW(201) & "mis par : " & d("Issuer"), False, 8, xlCenter).Font.Color = COLOR_GRAY_TEXT
    SetRowHeight ws, r, 13

    WriteReceiptCopy = r + 1
End Function

' Case de synthèse : libellé (ligne pRow) et montant (ligne pRow + 1) sur 2 colonnes
Private Sub PutSummaryBox(ByVal ws As Worksheet, ByVal pRow As Long, ByVal pCol As Long, ByVal pLabel As String, _
                          ByVal pAmount As Currency)
    With PutText(ws, pRow, pCol, pCol + 1, pLabel, True, 8, xlCenter)
        .WrapText = True
        .Interior.Color = COLOR_GRAY_FILL
    End With
    PutAmount ws, pRow + 1, pCol, pCol + 1, pAmount, True
    ws.Range(ws.Cells(pRow + 1, pCol), ws.Cells(pRow + 1, pCol + 1)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(pRow, pCol), ws.Cells(pRow + 1, pCol + 1)).Borders(xlEdgeLeft).Weight = xlThin
End Sub

' Libellé (colonnes 1-2) et valeur (colonnes 3-6)
Private Sub PutLabelValue(ByVal ws As Worksheet, ByVal pRow As Long, ByVal pLabel As String, ByVal pValue As String, _
                          ByVal pBold As Boolean)
    PutText ws, pRow, 1, 2, pLabel, False, 10, xlLeft
    PutText ws, pRow, 3, 6, pValue, pBold, 10, xlLeft
End Sub

'==============================================================================
' MISE EN PAGE : OUTILS COMMUNS
'==============================================================================

' Police, largeurs de colonnes (en points), quadrillage masqué
Private Sub PrepareSheet(ByVal ws As Worksheet)
    Dim i As Long

    ws.Cells.Font.Name = FONT_NAME
    ws.Cells.Font.Size = FONT_SIZE
    ws.Cells.VerticalAlignment = xlCenter
    For i = 1 To COL_COUNT
        SetColumnWidthPoints ws, i, COL_WIDTH_PT
    Next i
End Sub

' En-tête de l'entreprise sur les colonnes 1 à pLastCol à partir de la ligne pRow (4 lignes) ;
' logo facultatif à gauche (SOCIETE_LOGO). pLarge : grande taille (facture). Retourne la ligne suivante.
Private Function WriteCompanyHeader(ByVal ws As Worksheet, ByVal pRow As Long, ByVal pLastCol As Long, _
                                    ByVal pLarge As Boolean) As Long
    Dim lngFirstCol As Long
    Dim strLine As String

    lngFirstCol = 1
    If AddLogo(ws, pRow, IIf(pLarge, 52, 40)) Then lngFirstCol = 2

    PutText(ws, pRow, lngFirstCol, pLastCol, CompanySetting(PARAM_COMPANY_NAME), True, IIf(pLarge, 14, 12), xlLeft).Font.Color = COLOR_TITLE
    PutText ws, pRow + 1, lngFirstCol, pLastCol, CompanySetting(PARAM_COMPANY_ADDRESS), False, IIf(pLarge, 9, 8), xlLeft

    strLine = PrefixIfAny("T" & ChrW(233) & "l. : ", CompanySetting(PARAM_COMPANY_PHONE))
    If Len(CompanySetting(PARAM_COMPANY_EMAIL)) > 0 Then
        If Len(strLine) > 0 Then strLine = strLine & " - "
        strLine = strLine & CompanySetting(PARAM_COMPANY_EMAIL)
    End If
    PutText ws, pRow + 2, lngFirstCol, pLastCol, strLine, False, IIf(pLarge, 9, 8), xlLeft
    PutText(ws, pRow + 3, lngFirstCol, pLastCol, LegalIds(), False, IIf(pLarge, 8, 7), xlLeft).Font.Color = COLOR_GRAY_TEXT

    If pLarge Then SetRowHeight ws, pRow, 26 Else SetRowHeight ws, pRow, 18
    WriteCompanyHeader = pRow + 4
End Function

' Logo (paramètre SOCIETE_LOGO, voir modKamerixCompany.ImportLogo) en haut à gauche ; Faux s'il n'y en a pas ou s'il est illisible
Private Function AddLogo(ByVal ws As Worksheet, ByVal pRow As Long, ByVal pHeight As Double) As Boolean
    Dim strPath As String
    Dim shp As Object

    strPath = ResolveLogoPath(CompanySetting(PARAM_COMPANY_LOGO))      ' chemin relatif au dossier de Kamerix
    If Len(strPath) = 0 Then Exit Function
    If Not FileExists(strPath) Then Exit Function

    On Error GoTo Fail
    Set shp = ws.Shapes.AddPicture(strPath, False, True, ws.Cells(pRow, 1).Left + 2, ws.Cells(pRow, 1).Top + 2, -1, -1)
    shp.LockAspectRatio = -1                   ' msoTrue
    shp.Height = pHeight
    If shp.Width > COL_WIDTH_PT - 6 Then shp.Width = COL_WIDTH_PT - 6
    AddLogo = True
    Exit Function

Fail:
    AddLogo = False
End Function

' Texte dans une cellule ou une plage fusionnée ; retourne la plage
Private Function PutText(ByVal ws As Worksheet, ByVal pRow As Long, ByVal pFirstCol As Long, ByVal pLastCol As Long, _
                         ByVal pText As String, ByVal pBold As Boolean, ByVal pSize As Double, _
                         ByVal pAlign As Long) As Range
    Dim rng As Range

    Set rng = ws.Range(ws.Cells(pRow, pFirstCol), ws.Cells(pRow, pLastCol))
    If pLastCol > pFirstCol Then rng.Merge
    rng.NumberFormat = "@"
    rng.Cells(1, 1).Value = pText
    rng.Font.Bold = pBold
    rng.Font.Size = pSize
    rng.HorizontalAlignment = pAlign
    Set PutText = rng
End Function

' Montant au format de la devise (paramètres DEVISE et DECIMALES_DEVISE)
Private Function PutAmount(ByVal ws As Worksheet, ByVal pRow As Long, ByVal pFirstCol As Long, ByVal pLastCol As Long, _
                           ByVal pAmount As Currency, ByVal pBold As Boolean) As Range
    Dim rng As Range

    Set rng = ws.Range(ws.Cells(pRow, pFirstCol), ws.Cells(pRow, pLastCol))
    If pLastCol > pFirstCol Then rng.Merge
    rng.NumberFormat = AmountNumberFormat()
    rng.Cells(1, 1).Value = pAmount
    rng.Font.Bold = pBold
    rng.HorizontalAlignment = xlRight
    Set PutAmount = rng
End Function

Private Sub SetRowHeight(ByVal ws As Worksheet, ByVal pRow As Long, ByVal pHeight As Double)
    If pHeight > 409 Then pHeight = 409
    If pHeight < 0.75 Then pHeight = 0.75
    ws.Rows(pRow).RowHeight = pHeight
End Sub

' Cadre fin autour d'une plage
Private Sub BoxRange(ByVal rng As Range)
    Dim vEdges As Variant
    Dim i As Long

    vEdges = Array(xlEdgeLeft, xlEdgeTop, xlEdgeRight, xlEdgeBottom)
    For i = LBound(vEdges) To UBound(vEdges)
        With rng.Borders(vEdges(i))
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = COLOR_GRAY_TEXT
        End With
    Next i
End Sub

' Largeur de colonne en points : ColumnWidth est en caractères, on ajuste par proportion
Private Sub SetColumnWidthPoints(ByVal ws As Worksheet, ByVal pCol As Long, ByVal pPoints As Double)
    Dim i As Long

    ws.Columns(pCol).ColumnWidth = pPoints / 5.6
    For i = 1 To 3
        If ws.Columns(pCol).Width > 0 Then
            ws.Columns(pCol).ColumnWidth = ws.Columns(pCol).ColumnWidth * pPoints / ws.Columns(pCol).Width
        End If
    Next i
End Sub

' Mention en filigrane, centrée horizontalement, inclinée
Private Sub AddWatermark(ByVal ws As Worksheet, ByVal pText As String, ByVal pTop As Double, ByVal pHeight As Double)
    Dim shp As Object
    Dim dblWidth As Double

    dblWidth = COL_WIDTH_PT * COL_COUNT
    ' Constantes Office écrites en valeur : 1 = msoTextOrientationHorizontal, 0 = msoFalse,
    ' -1 = msoTrue, 3 = msoAnchorMiddle, 2 = msoAlignCenter
    Set shp = ws.Shapes.AddTextbox(1, 0, pTop, dblWidth, pHeight)
    shp.Fill.Visible = 0
    shp.Line.Visible = 0
    With shp.TextFrame2
        .WordWrap = 0
        .VerticalAnchor = 3
        .TextRange.Text = pText
        .TextRange.ParagraphFormat.Alignment = 2
        .TextRange.Font.Size = 72
        .TextRange.Font.Bold = -1
        .TextRange.Font.Fill.ForeColor.RGB = COLOR_RED
        .TextRange.Font.Fill.Transparency = 0.7
    End With
    shp.Rotation = -25
    shp.Placement = xlFreeFloating
End Sub

' Mise en page A4 portrait sur une seule page, centrée horizontalement.
' PrintCommunication groupe les réglages (beaucoup plus rapide) ; il est toujours rétabli.
Private Sub ApplyPageSetup(ByVal ws As Worksheet, ByVal pLastRow As Long, ByVal pTopCm As Double, _
                           ByVal pBottomCm As Double, ByVal pFooter As String)
    Dim lngErr As Long
    Dim strErr As String

    On Error Resume Next
    Application.PrintCommunication = False
    On Error GoTo EH

    With ws.PageSetup
        .PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(pLastRow, COL_COUNT)).Address
        .PaperSize = xlPaperA4
        .Orientation = xlPortrait
        .LeftMargin = Application.CentimetersToPoints(1.5)
        .RightMargin = Application.CentimetersToPoints(1.5)
        .TopMargin = Application.CentimetersToPoints(pTopCm)
        .BottomMargin = Application.CentimetersToPoints(pBottomCm)
        .HeaderMargin = Application.CentimetersToPoints(0.5)
        .FooterMargin = Application.CentimetersToPoints(0.6)
        .CenterHorizontally = True
        .CenterVertically = False
        .PrintGridlines = False
        .CenterHeader = ""
        .CenterFooter = pFooter
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With

    On Error Resume Next
    Application.PrintCommunication = True
    Exit Sub

EH:
    lngErr = Err.Number: strErr = Err.Description
    On Error Resume Next
    Application.PrintCommunication = True
    On Error GoTo 0
    Err.Raise ERR_DOCUMENT, "modKamerixDocuments.ApplyPageSetup", _
        "Mise en page impossible (erreur " & lngErr & " : " & strErr & "). " & _
        "V" & ChrW(233) & "rifiez qu'une imprimante est install" & ChrW(233) & "e dans Windows (au besoin " & _
        ChrW(171) & " Microsoft Print to PDF " & ChrW(187) & ")."
End Sub

' Pied de page de la facture : mentions libres, puis raison sociale et identifiants
Private Function FooterText() As String
    Dim strText As String

    strText = CompanySetting(PARAM_FOOTER)
    If Len(strText) > 0 Then strText = strText & vbLf
    strText = strText & CompanySetting(PARAM_COMPANY_NAME)
    If Len(LegalIds()) > 0 Then strText = strText & " - " & LegalIds()
    ' & est un code de mise en forme du pied de page ; un pied de page est limité à 255 caractères
    FooterText = "&8" & Left$(Replace(strText, "&", "&&"), 240)
End Function

' "NIU : ... - RCCM : ..." (vide si les deux paramètres sont vides)
Private Function LegalIds() As String
    Dim strText As String

    strText = PrefixIfAny("NIU : ", CompanySetting(PARAM_COMPANY_NIU))
    If Len(CompanySetting(PARAM_COMPANY_RCCM)) > 0 Then
        If Len(strText) > 0 Then strText = strText & " - "
        strText = strText & "RCCM : " & CompanySetting(PARAM_COMPANY_RCCM)
    End If
    LegalIds = strText
End Function

Private Function DuplicateText(ByVal pEdition As Long) As String
    DuplicateText = "DUPLICATA n" & ChrW(176) & " " & pEdition & " - " & Format$(Now, "dd/mm/yyyy hh:nn")
End Function

' Format de nombre Excel des montants : #,##0 "FCFA" (codes anglais de NumberFormat)
Private Function AmountNumberFormat() As String
    Dim lngDecimals As Long

    lngDecimals = CachedCurrencyDecimals()
    AmountNumberFormat = "#,##0"
    If lngDecimals > 0 Then AmountNumberFormat = AmountNumberFormat & "." & String$(lngDecimals, "0")
    AmountNumberFormat = AmountNumberFormat & " """ & Replace(CachedCurrencyCode(), """", "") & """"
End Function

' Montant en texte avec la devise : "2 000 000 FCFA" (séparateur de milliers de Windows)
Private Function FormatAmountText(ByVal pAmount As Currency) As String
    Dim lngDecimals As Long

    lngDecimals = CachedCurrencyDecimals()
    If lngDecimals > 0 Then
        FormatAmountText = Format$(pAmount, "#,##0." & String$(lngDecimals, "0"))
    Else
        FormatAmountText = Format$(pAmount, "#,##0")
    End If
    FormatAmountText = FormatAmountText & " " & CachedCurrencyCode()
End Function

' Taux affiché : 19,25 ou 19 (sans séparateur final)
Private Function RateText(ByVal pRate As Double) As String
    If pRate = Int(pRate) Then RateText = Format$(pRate, "0") Else RateText = Format$(pRate, "0.00")
End Function

Private Function PrefixIfAny(ByVal pPrefix As String, ByVal pValue As String) As String
    If Len(pValue) > 0 Then PrefixIfAny = pPrefix & pValue
End Function

'==============================================================================
' CACHE DES PARAMETRES (PROCEDURES INTERNES)
'==============================================================================

' Tous les paramètres (clé en majuscules -> valeur texte, "" si vide)
Private Function SettingsMap() As Object
    Dim dic As Object
    Dim rs As ADODB.Recordset

    Set dic = CreateObject("Scripting.Dictionary")
    Set rs = QueryRecordset("SELECT [Cle], [Valeur] FROM [Parametre]")
    Do While Not rs.EOF
        dic(UCase$(CStr(rs.Fields("Cle").Value))) = TextOf(rs.Fields("Valeur").Value)
        rs.MoveNext
    Loop
    rs.Close
    Set SettingsMap = dic
End Function

Private Sub BeginSettingsCache()
    mstrCurrencyCode = CurrencyCode()
    mlngCurrencyDecimals = CurrencyDecimals()
    Set mdicSettings = SettingsMap()
End Sub

Private Sub EndSettingsCache()
    Set mdicSettings = Nothing
    mstrCurrencyCode = ""
    mlngCurrencyDecimals = 0
End Sub

Private Function CachedCurrencyCode() As String
    If Len(mstrCurrencyCode) > 0 Then CachedCurrencyCode = mstrCurrencyCode Else CachedCurrencyCode = CurrencyCode()
End Function

Private Function CachedCurrencyDecimals() As Long
    If mdicSettings Is Nothing Then CachedCurrencyDecimals = CurrencyDecimals() Else CachedCurrencyDecimals = mlngCurrencyDecimals
End Function

'==============================================================================
' FICHIERS, DROITS ET JOURNAL (PROCEDURES INTERNES)
'==============================================================================

Private Sub RequireDocumentRead(ByVal pDocType As String)
    Select Case pDocType
        Case DOC_INVOICE
            RequireRead MOD_INVOICES
        Case DOC_RECEIPT, DOC_REFUND
            RequireRead MOD_PAYMENTS
        Case Else
            Err.Raise ERR_DOCUMENT, "modKamerixDocuments", "Type de document inconnu : " & pDocType & "."
    End Select
End Sub

Private Sub LogEdition(ByVal pDocType As String, ByVal pDocId As Long, ByVal pAction As String, ByVal pDuplicate As Boolean)
    ExecuteInsert "INSERT INTO [JournalDocument] ([TypeDocument], [DocumentID], [Action], [Duplicata], [UtilisateurID], [DateHeure]) " & _
                  "VALUES (?, ?, ?, ?, ?, ?)", pDocType, pDocId, pAction, pDuplicate, CurrentUserId(), Now
End Sub

' Dossier du type de document pour l'année de la pièce (créé s'il n'existe pas)
Private Function DocumentFolder(ByVal pDocType As String, ByVal pDate As Date) As String
    Dim strRoot As String
    Dim strSub As String

    If Len(mstrTestRoot) > 0 Then
        strRoot = mstrTestRoot
    Else
        If Len(ThisWorkbook.Path) = 0 Then
            Err.Raise ERR_DOCUMENT, "modKamerixDocuments.DocumentFolder", "Enregistrez d'abord le classeur Kamerix."
        End If
        strRoot = ThisWorkbook.Path & Application.PathSeparator & "Documents"
    End If

    Select Case pDocType
        Case DOC_INVOICE: strSub = "Factures"
        Case DOC_RECEIPT: strSub = "Recus"
        Case DOC_REFUND:  strSub = "Remboursements"
    End Select

    DocumentFolder = strRoot & Application.PathSeparator & CStr(Year(pDate)) & Application.PathSeparator & strSub
    EnsureFolder DocumentFolder
End Function

' Nom du PDF sans extension : FAC-2026-00001 (original) ou FAC-2026-00001_DUP2
Private Function PdfBaseName(ByVal pNumber As String, ByVal pEdition As Long) As String
    PdfBaseName = pNumber
    If pEdition > 1 Then PdfBaseName = PdfBaseName & "_DUP" & pEdition
End Function

' Chemin libre : jamais d'écrasement d'un fichier existant
Private Function UniqueFilePath(ByVal pFolder As String, ByVal pBaseName As String) As String
    Dim strPath As String
    Dim i As Long

    strPath = pFolder & Application.PathSeparator & pBaseName & ".pdf"
    i = 1
    Do While FileExists(strPath)
        i = i + 1
        strPath = pFolder & Application.PathSeparator & pBaseName & " (" & i & ").pdf"
    Loop
    UniqueFilePath = strPath
End Function

' Crée le dossier et ses parents manquants
Private Sub EnsureFolder(ByVal pFolder As String)
    Dim lngPos As Long

    If FolderExists(pFolder) Then Exit Sub
    If Right$(pFolder, 1) = ":" Then Exit Sub         ' racine d'un lecteur (C:)
    lngPos = InStrRev(pFolder, Application.PathSeparator)
    If lngPos > 1 Then EnsureFolder Left$(pFolder, lngPos - 1)
    MkDir pFolder
End Sub

Private Function FolderExists(ByVal pFolder As String) As Boolean
    On Error Resume Next
    FolderExists = ((GetAttr(pFolder) And vbDirectory) = vbDirectory)
    On Error GoTo 0
End Function

Private Function FileExists(ByVal pPath As String) As Boolean
    On Error Resume Next
    FileExists = ((GetAttr(pPath) And vbDirectory) = 0)
    If Err.Number <> 0 Then FileExists = False
    On Error GoTo 0
End Function

Private Function RecordCount(ByVal rs As ADODB.Recordset) As Long
    If rs.EOF Then Exit Function
    rs.MoveFirst
    Do While Not rs.EOF
        RecordCount = RecordCount + 1
        rs.MoveNext
    Loop
    rs.MoveFirst
End Function

Private Function TextOf(ByVal pValue As Variant) As String
    If Not IsNull(pValue) And Not IsEmpty(pValue) Then TextOf = Trim$(CStr(pValue))
End Function
