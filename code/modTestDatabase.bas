Attribute VB_Name = "modTestDatabase"
Option Explicit
Option Private Module
'==============================================================================
' KAMERIX - Tests de la base Kamerixdb.accdb (ouverture protegee par mot de passe)
'------------------------------------------------------------------------------
' Lancer TestDatabaseOpening (F5) : execute 8 controles et affiche un rapport.
'   1. Le fichier de la base existe
'   2. Ouverture avec le mot de passe
'   3. Nombre de tables conforme
'   4. Donnees de reference et version du schema
'   5. Connexion en lecture seule : lecture OK, ecriture refusee
'   6. Ecriture reelle (annulee ensuite par ROLLBACK : aucune donnee laissee)
'   7. Ouverture SANS mot de passe : doit etre refusee
'   8. Ouverture avec un MAUVAIS mot de passe : doit etre refusee
' Les tests 7 et 8 prouvent que la protection est reellement active.
' Aucun de ces tests ne modifie durablement la base.
'
' Dependances : modKamerixDb (OpenConnection, GetDatabasePath, DatabaseExists),
'               modKamerixSchema (ExpectedTables, SCHEMA_VERSION).
' Le fournisseur est repete ici volontairement pour garder les tests 7 et 8
' independants du code teste.
'==============================================================================

Private Const PROVIDER As String = "Microsoft.ACE.OLEDB.12.0"
Private Const AD_STATE_CLOSED As Long = 0
Private Const AD_STATE_OPEN As Long = 1
Private Const WRONG_PASSWORD As String = "mot_de_passe_incorrect"

'==============================================================================
' PROCEDURE PRINCIPALE
'==============================================================================
Public Sub TestDatabaseOpening()
    Dim strReport As String
    Dim lngFailures As Long
    Dim strDetail As String
    Dim blnOk As Boolean

    On Error GoTo EH

    strReport = "Base : " & GetDatabasePath() & vbCrLf & vbCrLf

    blnOk = Test_FileExists(strDetail)
    AddResult strReport, lngFailures, "1. Fichier de la base present", blnOk, strDetail
    If Not blnOk Then GoTo FINISH

    blnOk = Test_OpenWithPassword(strDetail)
    AddResult strReport, lngFailures, "2. Ouverture avec le mot de passe", blnOk, strDetail
    If Not blnOk Then GoTo FINISH

    blnOk = Test_TableCount(strDetail)
    AddResult strReport, lngFailures, "3. Nombre de tables", blnOk, strDetail

    blnOk = Test_ReferenceData(strDetail)
    AddResult strReport, lngFailures, "4. Donnees de reference et version", blnOk, strDetail

    blnOk = Test_ReadOnly(strDetail)
    AddResult strReport, lngFailures, "5. Lecture seule (ecriture refusee)", blnOk, strDetail

    blnOk = Test_WriteWithRollback(strDetail)
    AddResult strReport, lngFailures, "6. Ecriture annulee par ROLLBACK", blnOk, strDetail

    blnOk = Test_RejectWithoutPassword(strDetail)
    AddResult strReport, lngFailures, "7. Refus sans mot de passe", blnOk, strDetail

    blnOk = Test_RejectWrongPassword(strDetail)
    AddResult strReport, lngFailures, "8. Refus avec mauvais mot de passe", blnOk, strDetail

FINISH:
    If lngFailures = 0 Then
        strReport = strReport & vbCrLf & "RESULTAT : tous les tests sont passes."
        Debug.Print strReport
        MsgBox strReport, vbInformation, "Kamerix - Tests de la base"
    Else
        strReport = strReport & vbCrLf & "RESULTAT : " & lngFailures & " test(s) en echec."
        Debug.Print strReport
        MsgBox strReport, vbExclamation, "Kamerix - Tests de la base"
    End If
    Exit Sub

EH:
    MsgBox "Erreur inattendue dans les tests (" & Err.Number & ") :" & vbCrLf & Err.Description, _
           vbCritical, "Kamerix - Tests de la base"
End Sub

'==============================================================================
' EXEMPLE D'UTILISATION (modele a suivre dans le reste de l'application)
' Affiche les specialites dans la fenetre Execution (Ctrl+G).
'==============================================================================
Public Sub ExampleReadSpecialties()
    Dim Adocn As Object
    Dim AdoRs As Object
    Dim lngErr As Long
    Dim strErr As String

    On Error GoTo EH

    Set Adocn = OpenConnection(True)
    Set AdoRs = Adocn.Execute("SELECT [SpecialiteID], [Libelle] FROM [Specialite] ORDER BY [Libelle]")
    Do While Not AdoRs.EOF
        Debug.Print AdoRs.Fields("SpecialiteID").Value & " - " & AdoRs.Fields("Libelle").Value
        AdoRs.MoveNext
    Loop
    AdoRs.Close
    Adocn.Close
    Exit Sub
EH:
    lngErr = Err.Number
    strErr = Err.Description
    On Error Resume Next
    If Not AdoRs Is Nothing Then AdoRs.Close
    CloseQuietly Adocn
    On Error GoTo 0
    MsgBox "Echec (" & lngErr & ") : " & strErr, vbCritical, "Kamerix"
End Sub

'==============================================================================
' LES 8 TESTS (chacun retourne True si OK et renseigne strDetail)
'==============================================================================

Private Function Test_FileExists(ByRef strDetail As String) As Boolean
    On Error GoTo EH

    Test_FileExists = DatabaseExists()
    If Test_FileExists Then
        strDetail = "trouve"
    Else
        strDetail = "fichier absent : lancer TestDatabaseCreation d'abord"
    End If
    Exit Function
EH:
    strDetail = ErrText()
End Function

Private Function Test_OpenWithPassword(ByRef strDetail As String) As Boolean
    Dim Adocn As Object

    On Error GoTo EH

    Set Adocn = OpenConnection()
    If Adocn.State = AD_STATE_OPEN Then
        Test_OpenWithPassword = True
        strDetail = "connexion ouverte"
    Else
        strDetail = "connexion non ouverte"
    End If
    CloseQuietly Adocn
    Exit Function
EH:
    strDetail = ErrText()
    CloseQuietly Adocn
End Function

Private Function Test_TableCount(ByRef strDetail As String) As Boolean
    Dim Adocn As Object
    Dim cat As Object
    Dim tbl As Object
    Dim vNames As Variant
    Dim lngCount As Long
    Dim lngExpected As Long

    On Error GoTo EH

    Set Adocn = OpenConnection(True)
    Set cat = CreateObject("ADOX.Catalog")
    Set cat.ActiveConnection = Adocn

    For Each tbl In cat.Tables
        If tbl.Type = "TABLE" Then lngCount = lngCount + 1
    Next tbl

    vNames = ExpectedTables()
    lngExpected = UBound(vNames) - LBound(vNames) + 1

    strDetail = lngCount & " table(s), " & lngExpected & " attendue(s)"
    Test_TableCount = (lngCount = lngExpected)

    Set cat = Nothing
    CloseQuietly Adocn
    Exit Function

EH:
    strDetail = ErrText()
    Set cat = Nothing
    CloseQuietly Adocn
End Function

Private Function Test_ReferenceData(ByRef strDetail As String) As Boolean
    Dim Adocn As Object
    Dim lngSpec As Long
    Dim lngCat As Long
    Dim lngVer As Long

    On Error GoTo EH

    Set Adocn = OpenConnection(True)
    lngSpec = ScalarLong(Adocn, "SELECT COUNT(*) FROM [Specialite]")
    lngCat = ScalarLong(Adocn, "SELECT COUNT(*) FROM [CategorieCharge]")
    lngVer = ScalarLong(Adocn, "SELECT MAX([Numero]) FROM [VersionSchema]")

    strDetail = lngSpec & " specialites, " & lngCat & " categories de charges, schema v" & lngVer
    Test_ReferenceData = (lngSpec >= 6 And lngCat >= 6 And lngVer = SCHEMA_VERSION)

    CloseQuietly Adocn
    Exit Function

EH:
    strDetail = ErrText()
    CloseQuietly Adocn
End Function

Private Function Test_ReadOnly(ByRef strDetail As String) As Boolean
    Dim Adocn As Object
    Dim lngRows As Long
    Dim blnRefused As Boolean

    On Error GoTo EH

    Set Adocn = OpenConnection(True)
    lngRows = ScalarLong(Adocn, "SELECT COUNT(*) FROM [Marque]")

    On Error Resume Next
    Adocn.Execute "INSERT INTO [Marque] ([Nom]) VALUES ('__TEST_LECTURE_SEULE__')", , 129
    blnRefused = (Err.Number <> 0)
    Err.Clear
    On Error GoTo EH

    If blnRefused Then
        strDetail = "lecture OK, ecriture refusee comme attendu"
    Else
        strDetail = "ATTENTION : une ecriture a ete acceptee en lecture seule (ligne '__TEST_LECTURE_SEULE__' dans Marque a supprimer)"
    End If
    Test_ReadOnly = blnRefused

    CloseQuietly Adocn
    Exit Function

EH:
    strDetail = ErrText()
    CloseQuietly Adocn
End Function

Private Function Test_WriteWithRollback(ByRef strDetail As String) As Boolean
    Dim Adocn As Object
    Dim blnInTrans As Boolean
    Dim lngDuring As Long
    Dim lngAfter As Long
    Dim strCount As String

    On Error GoTo EH

    strCount = "SELECT COUNT(*) FROM [Marque] WHERE [Nom] = '__TEST_KAMERIX__'"

    Set Adocn = OpenConnection()
    Adocn.BeginTrans
    blnInTrans = True

    Adocn.Execute "INSERT INTO [Marque] ([Nom]) VALUES ('__TEST_KAMERIX__')", , 129
    lngDuring = ScalarLong(Adocn, strCount)

    Adocn.RollbackTrans
    blnInTrans = False
    lngAfter = ScalarLong(Adocn, strCount)

    strDetail = "pendant la transaction : " & lngDuring & " ligne, apres ROLLBACK : " & lngAfter & " ligne"
    Test_WriteWithRollback = (lngDuring = 1 And lngAfter = 0)

    CloseQuietly Adocn
    Exit Function

EH:
    strDetail = ErrText()
    On Error Resume Next
    If blnInTrans Then Adocn.RollbackTrans
    On Error GoTo 0
    CloseQuietly Adocn
End Function

Private Function Test_RejectWithoutPassword(ByRef strDetail As String) As Boolean
    Dim strErr As String

    On Error GoTo EH

    If TryOpen("Provider=" & PROVIDER & ";Data Source=" & GetDatabasePath() & ";Mode=Read;", strErr) Then
        strDetail = "ATTENTION : la base s'ouvre sans mot de passe, elle n'est pas protegee"
    Else
        Test_RejectWithoutPassword = True
        strDetail = "acces refuse comme attendu"
    End If
    Exit Function

EH:
    strDetail = ErrText()
End Function

Private Function Test_RejectWrongPassword(ByRef strDetail As String) As Boolean
    Dim strErr As String

    On Error GoTo EH

    If TryOpen("Provider=" & PROVIDER & ";Data Source=" & GetDatabasePath() & _
               ";Jet OLEDB:Database Password=" & WRONG_PASSWORD & ";Mode=Read;", strErr) Then
        strDetail = "ATTENTION : la base s'ouvre avec un mauvais mot de passe"
    Else
        Test_RejectWrongPassword = True
        strDetail = "acces refuse comme attendu"
    End If
    Exit Function

EH:
    strDetail = ErrText()
End Function

'==============================================================================
' OUTILS INTERNES
'==============================================================================

Private Sub AddResult(ByRef strReport As String, ByRef lngFailures As Long, _
                      ByVal strTitle As String, ByVal blnOk As Boolean, ByVal strDetail As String)
    If blnOk Then
        strReport = strReport & "[OK]      " & strTitle
    Else
        strReport = strReport & "[ECHEC]   " & strTitle
        lngFailures = lngFailures + 1
    End If
    If Len(strDetail) > 0 Then strReport = strReport & vbCrLf & "              " & strDetail
    strReport = strReport & vbCrLf
End Sub

' Tente d'ouvrir une connexion. Retourne True si elle s'ouvre.
Private Function TryOpen(ByVal strConnection As String, ByRef strErr As String) As Boolean
    Dim Adocn As Object

    On Error Resume Next
    Set Adocn = CreateObject("ADODB.Connection")
    Adocn.Open strConnection
    If Err.Number = 0 Then
        TryOpen = True
    Else
        strErr = Err.Description
    End If
    Err.Clear
    CloseQuietly Adocn
    Set Adocn = Nothing
    On Error GoTo 0
End Function

' Retourne la premiere colonne de la premiere ligne en Long (0 si vide ou Null)
Private Function ScalarLong(ByVal Adocn As Object, ByVal strSql As String) As Long
    Dim AdoRs As Object

    Set AdoRs = Adocn.Execute(strSql)
    If Not AdoRs.EOF Then
        If Not IsNull(AdoRs.Fields(0).Value) Then ScalarLong = CLng(AdoRs.Fields(0).Value)
    End If
    AdoRs.Close
    Set AdoRs = Nothing
End Function

Private Function ErrText() As String
    ErrText = "erreur " & Err.Number & " : " & Replace(Err.Description, vbCrLf, " ")
End Function

Private Sub CloseQuietly(ByVal Adocn As Object)
    On Error Resume Next
    If Not Adocn Is Nothing Then
        If Adocn.State <> AD_STATE_CLOSED Then Adocn.Close
    End If
    On Error GoTo 0
End Sub
