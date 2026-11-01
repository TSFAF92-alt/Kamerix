Attribute VB_Name = "modTestListFill"
Option Explicit
Option Private Module

'==============================================================================
' KAMERIX - Tests du lot 5.4a : conversion d'un recordset pour les ListBox
'------------------------------------------------------------------------------
' A lancer (F5 sur TestListFill). Lit des donnees de reference existantes ;
' n'ecrit rien dans la base.
'==============================================================================

Public Sub TestListFill()
    Dim lngOk As Long
    Dim lngTotal As Long
    Dim strReport As String
    Dim vData As Variant
    Dim lngRoles As Long

    On Error GoTo EH

    lngRoles = CLng(QueryValue("SELECT COUNT(*) FROM [Role]"))
    vData = RecordsetToColumns(QueryRecordset("SELECT [RoleID], [Nom] FROM [Role] ORDER BY [RoleID]"), Array("RoleID", "Nom"))
    CheckTrue "Tableau obtenu", IsArray(vData), lngOk, lngTotal, strReport
    CheckEq "Orientation (colonne, ligne) : 2 colonnes", UBound(vData, 1) + 1, 2, lngOk, lngTotal, strReport
    CheckEq "Une ligne par role", UBound(vData, 2) + 1, lngRoles, lngOk, lngTotal, strReport
    CheckEq "Premier role : Administrateur", vData(1, 0), "Administrateur", lngOk, lngTotal, strReport

    vData = RecordsetToColumns(QueryRecordset("SELECT [RoleID], [Nom] FROM [Role] WHERE [RoleID] = ?", -1), Array("RoleID", "Nom"))
    CheckTrue "Recordset vide : Empty (aucun appel a GetRows)", IsEmpty(vData), lngOk, lngTotal, strReport

    vData = RecordsetToColumns(QueryRecordset("SELECT [VersionID], [Numero], [DateApplication], Null AS Vide FROM [VersionSchema]"), _
                               Array("Numero", "DateApplication", "Vide"), Array("", "yyyy", ""))
    CheckEq "Format applique (annee de la date)", vData(1, 0), CStr(Year(CDate(QueryValue( _
        "SELECT MIN([DateApplication]) FROM [VersionSchema]")))), lngOk, lngTotal, strReport
    CheckEq "Null remplace par une chaine vide", "[" & vData(2, 0) & "]", "[]", lngOk, lngTotal, strReport
    CheckEq "Sans format : valeur conservee", vData(0, 0), SCHEMA_VERSION, lngOk, lngTotal, strReport

    Debug.Print lngOk & " tests reussis sur " & lngTotal & "." & vbCrLf & strReport   ' rapport complet (Ctrl+G)
    MsgBox lngOk & " tests reussis sur " & lngTotal & "." & vbCrLf & vbCrLf & strReport, _
        IIf(lngOk = lngTotal, vbInformation, vbExclamation), "Kamerix - tests des listes"
    Exit Sub

EH:
    MsgBox "Echec (erreur " & Err.Number & ") :" & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
           "Resultats deja obtenus :" & vbCrLf & strReport, vbCritical, "Kamerix - tests des listes"
End Sub

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
