Attribute VB_Name = "modKamerixNumbering"
Option Explicit
Option Private Module

'==============================================================================
' KAMERIX - Numerotation sequentielle des pieces (table Compteur)
'------------------------------------------------------------------------------
' Format      : TYPE-ANNEE-NNNNN, par exemple FAC-2026-00001. Un compteur par
'               type de piece et par annee : il repart a 1 chaque annee.
' Types       : DOC_INVOICE (FAC), DOC_RECEIPT (REC), DOC_REFUND (RMB).
' Transaction : NextDocumentNumber doit etre appele DANS la transaction qui
'               enregistre la piece. Si la piece n'est pas enregistree
'               (RollbackTransaction), le compteur revient en arriere : aucun
'               numero n'est perdu ni attribue deux fois.
' Dependances : modKamerixSql, modKamerixConst.
'==============================================================================

' Prochain numero d'une piece. pDate : date de la piece (annee du compteur), aujourd'hui par defaut.
Public Function NextDocumentNumber(ByVal pDocType As String, Optional ByVal pDate As Variant) As String
    Dim lngYear As Long
    Dim lngValue As Long

    If Not TransactionActive() Then
        Err.Raise ERR_NUMBERING, "modKamerixNumbering.NextDocumentNumber", _
            "La num" & ChrW(233) & "rotation doit se faire dans la transaction qui enregistre la pi" & ChrW(232) & "ce."
    End If
    If Not IsDocumentType(pDocType) Then
        Err.Raise ERR_NUMBERING, "modKamerixNumbering.NextDocumentNumber", _
            "Type de pi" & ChrW(232) & "ce inconnu : " & pDocType
    End If

    If IsMissing(pDate) Then lngYear = Year(Date) Else lngYear = Year(CDate(pDate))

    If ExecuteNonQuery("UPDATE [Compteur] SET [DerniereValeur] = [DerniereValeur] + 1 " & _
                       "WHERE [TypeDocument] = ? AND [Annee] = ?", pDocType, lngYear) = 0 Then
        ExecuteInsert "INSERT INTO [Compteur] ([TypeDocument], [Annee], [DerniereValeur]) VALUES (?, ?, ?)", _
                      pDocType, lngYear, 1
    End If

    lngValue = CLng(QueryValue("SELECT [DerniereValeur] FROM [Compteur] WHERE [TypeDocument] = ? AND [Annee] = ?", _
                               pDocType, lngYear))
    NextDocumentNumber = FormatDocumentNumber(pDocType, lngYear, lngValue)
End Function

' Mise en forme d'un numero : FAC-2026-00001
Public Function FormatDocumentNumber(ByVal pDocType As String, ByVal pYear As Long, ByVal pValue As Long) As String
    FormatDocumentNumber = pDocType & "-" & CStr(pYear) & "-" & Format$(pValue, "00000")
End Function

Private Function IsDocumentType(ByVal pDocType As String) As Boolean
    Select Case pDocType
        Case DOC_INVOICE, DOC_RECEIPT, DOC_REFUND
            IsDocumentType = True
    End Select
End Function
