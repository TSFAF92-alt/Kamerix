Attribute VB_Name = "modTestSecuritySeed"
Option Explicit
Option Private Module

'==============================================================================
' KAMERIX - Test des donnees de securite du schema v5
' A lancer (F5) apres suppression puis recreation de la base de developpement.
'==============================================================================

Public Sub TestSecuritySeed()
    Dim Adocn As ADODB.Connection
    Dim lngOk As Long
    Dim lngTotal As Long
    Dim strReport As String

    On Error GoTo EH

    Set Adocn = OpenConnection(True)

    Check Adocn, "Version du schema = 5", "SELECT MAX([Numero]) FROM [VersionSchema]", 5, lngOk, lngTotal, strReport
    Check Adocn, "4 roles", "SELECT COUNT(*) FROM [Role]", 4, lngOk, lngTotal, strReport
    Check Adocn, "9 modules", "SELECT COUNT(*) FROM [Module]", 9, lngOk, lngTotal, strReport
    Check Adocn, "36 permissions (4 roles x 9 modules)", "SELECT COUNT(*) FROM [Permission]", 36, lngOk, lngTotal, strReport
    Check Adocn, "Suppression accordee a 9 lignes (Administrateur seul)", _
        "SELECT COUNT(*) FROM [Permission] WHERE [Suppression] = True", 9, lngOk, lngTotal, strReport
    Check Adocn, "Administration inaccessible sauf Administrateur (3 lignes sans lecture)", _
        "SELECT COUNT(*) FROM [Permission] AS p INNER JOIN [Module] AS m ON p.[ModuleID] = m.[ModuleID] " & _
        "WHERE m.[Nom] = 'Administration' AND p.[Lecture] = False", 3, lngOk, lngTotal, strReport
    Check Adocn, "15 droits speciaux (6 + 6 + 3)", "SELECT COUNT(*) FROM [DroitSpecial]", 15, lngOk, lngTotal, strReport
    Check Adocn, "Vendeur sans droit special", _
        "SELECT COUNT(*) FROM [DroitSpecial] AS d INNER JOIN [Role] AS r ON d.[RoleID] = r.[RoleID] " & _
        "WHERE r.[Nom] = 'Vendeur'", 0, lngOk, lngTotal, strReport
    Check Adocn, "Aucun utilisateur a la creation (premier admin cree au 1er lancement)", _
        "SELECT COUNT(*) FROM [Utilisateur]", 0, lngOk, lngTotal, strReport

    Adocn.Close
    Set Adocn = Nothing

    Debug.Print lngOk & " tests reussis sur " & lngTotal & "." & vbCrLf & strReport   ' rapport complet (Ctrl+G)
    MsgBox lngOk & " tests reussis sur " & lngTotal & "." & vbCrLf & vbCrLf & strReport, _
        IIf(lngOk = lngTotal, vbInformation, vbExclamation), "Kamerix - donnees de securite"
    Exit Sub

EH:
    On Error Resume Next
    If Not Adocn Is Nothing Then
        If Adocn.State <> adStateClosed Then Adocn.Close
    End If
    On Error GoTo 0
    MsgBox "Echec (erreur " & Err.Number & ") :" & vbCrLf & vbCrLf & Err.Description, vbCritical, "Kamerix"
End Sub

Private Sub Check(ByVal Adocn As ADODB.Connection, ByVal strLabel As String, ByVal strSql As String, _
                  ByVal lngExpected As Long, ByRef lngOk As Long, ByRef lngTotal As Long, _
                  ByRef strReport As String)
    Dim AdoRs As ADODB.Recordset
    Dim lngActual As Long

    lngTotal = lngTotal + 1
    Set AdoRs = Adocn.Execute(strSql)
    If Not AdoRs.EOF Then
        If Not IsNull(AdoRs.Fields(0).Value) Then lngActual = CLng(AdoRs.Fields(0).Value)
    End If
    AdoRs.Close
    Set AdoRs = Nothing

    If lngActual = lngExpected Then
        lngOk = lngOk + 1
        strReport = strReport & "OK     " & strLabel & vbCrLf
    Else
        strReport = strReport & "ECHEC  " & strLabel & " (attendu " & lngExpected & ", obtenu " & lngActual & ")" & vbCrLf
    End If
End Sub
