Attribute VB_Name = "modKamerixListFill"
Option Explicit
Option Private Module

'==============================================================================
' KAMERIX - Remplissage des ListBox et ComboBox directement depuis la base
'------------------------------------------------------------------------------
' Principe    : les donnees sont lues dans Access (recordset) et chargees en une
'               seule affectation dans le controle, sans feuille intermediaire :
'                   lstItems.Column = AdoRs.GetRows(...)   (tableau colonne, ligne)
'               La propriete Column attend la meme orientation que GetRows ;
'               la propriete List attendrait l'inverse (ligne, colonne).
' Usage       : FillListControl lstItems, BrandList(), Array("MarqueID", "Nom")
'               FillListControl lstSales, rs, Array("Numero", "Date", "Prix"), _
'                               Array("", "dd/mm/yyyy", "#,##0")
' Precautions : - un recordset vide vide le controle (GetRows echouerait) ;
'               - les valeurs Null deviennent des chaines vides (une ListBox
'                 refuse un tableau qui contient Null) ;
'               - les en-tetes de colonnes sont des libelles poses au-dessus de
'                 la liste (ColumnHeads ne fonctionne qu'avec RowSource).
' Dependances : ADODB.
'==============================================================================

' Charge les champs pFields du recordset dans une ListBox ou une ComboBox (ferme le recordset).
'   pFormats : formats facultatifs par champ ("" = aucun), par exemple "dd/mm/yyyy" ou "#,##0"
' Retourne le nombre de lignes chargees.
Public Function FillListControl(ByVal pControl As Object, ByVal pRs As ADODB.Recordset, ByVal pFields As Variant, _
                                Optional ByVal pFormats As Variant) As Long
    Dim vData As Variant

    vData = RecordsetToColumns(pRs, pFields, pFormats)
    pControl.Clear
    pControl.ColumnCount = UBound(pFields) - LBound(pFields) + 1
    If IsEmpty(vData) Then Exit Function

    pControl.Column = vData
    FillListControl = UBound(vData, 2) + 1
End Function

' Tableau (colonne, ligne), indices a partir de 0, des champs pFields du recordset,
' lu par GetRows ; Null remplaces par "" et formats appliques. Empty si le recordset
' est vide (GetRows leverait une erreur). Ferme le recordset.
' (Separe de FillListControl pour etre teste sans formulaire.)
Public Function RecordsetToColumns(ByVal pRs As ADODB.Recordset, ByVal pFields As Variant, _
                                   Optional ByVal pFormats As Variant) As Variant
    Dim vData As Variant
    Dim i As Long
    Dim j As Long
    Dim strFormat As String

    If pRs.EOF Then
        pRs.Close
        Exit Function
    End If

    vData = pRs.GetRows(Rows:=adGetRowsRest, Fields:=pFields)
    pRs.Close

    For j = 0 To UBound(vData, 1)
        strFormat = ""
        If Not IsMissing(pFormats) Then strFormat = CStr(pFormats(LBound(pFormats) + j))
        For i = 0 To UBound(vData, 2)
            If IsNull(vData(j, i)) Then
                vData(j, i) = ""
            ElseIf Len(strFormat) > 0 Then
                vData(j, i) = Format$(vData(j, i), strFormat)
            End If
        Next i
    Next j

    RecordsetToColumns = vData
End Function

' Selectionne la ligne dont la premiere colonne (identifiant) vaut pId ; aucune si absente.
' Retourne Vrai si la ligne a ete trouvee.
Public Function SelectById(ByVal pControl As Object, ByVal pId As Long) As Boolean
    Dim i As Long

    For i = 0 To pControl.ListCount - 1
        If CLng(pControl.List(i, 0)) = pId Then
            pControl.ListIndex = i
            SelectById = True
            Exit Function
        End If
    Next i
    pControl.ListIndex = -1
End Function

' Identifiant (premiere colonne) de la ligne selectionnee ; 0 si aucune
Public Function SelectedIdOf(ByVal pControl As Object) As Long
    If pControl.ListIndex >= 0 Then SelectedIdOf = CLng(pControl.List(pControl.ListIndex, 0))
End Function

' Ajoute une ligne "(aucun)" d'identifiant 0 en tete d'une liste deroulante a 2 colonnes
Public Sub AddNoneItem(ByVal pControl As Object, Optional ByVal pCaption As String = "(aucun)")
    pControl.AddItem "0", 0
    pControl.List(0, 1) = pCaption
End Sub
