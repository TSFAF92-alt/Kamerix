Attribute VB_Name = "modTestSql"
Option Explicit
Option Private Module
'==============================================================================
' Tests de la couche SQL generique (modKamerixSql)
' Lancer TestSql avec F5. La base de developpement doit exister.
' Toutes les ecritures de test sont annulees par un Rollback, sauf un cas
' Commit qui est immediatement nettoye par un DELETE.
' Les liens entre tables (cles etrangeres) sont actifs : les lignes de test
' creent leurs parents (Role, Module, CategorieCharge, Utilisateur).
'==============================================================================
Private mlngPassed As Long
Private mlngFailed As Long

Private Sub Check(ByVal strName As String, ByVal blnCondition As Boolean)
    If blnCondition Then
        mlngPassed = mlngPassed + 1
        Debug.Print "OK   - " & strName
    Else
        mlngFailed = mlngFailed + 1
        Debug.Print "ECHEC - " & strName
    End If
End Sub

' Texte d'une valeur lue (jamais d'erreur sur Null)
Private Function TextOf(ByVal vValue As Variant) As String
    If IsNull(vValue) Then
        TextOf = "<Null>"
    Else
        TextOf = CStr(vValue)
    End If
End Function

Private Function IsTrueValue(ByVal vValue As Variant) As Boolean
    If Not IsNull(vValue) Then IsTrueValue = CBool(vValue)
End Function

Private Function IsFalseValue(ByVal vValue As Variant) As Boolean
    If Not IsNull(vValue) Then IsFalseValue = Not CBool(vValue)
End Function

Public Sub TestSql()
    Dim lngBefore As Long
    Dim lngId As Long
    Dim lngRows As Long
    Dim vVal As Variant
    Dim AdoRs As ADODB.Recordset
    Dim strText As String
    Dim dtTest As Date
    Dim blnOk As Boolean
    Dim lngRole As Long
    Dim lngModule As Long
    Dim lngCategory As Long
    Dim lngUser As Long

    mlngPassed = 0
    mlngFailed = 0
    Debug.Print "=== Tests couche SQL generique ==="

    On Error GoTo EH

    ' --- Connexion et lecture simple
    lngBefore = CLng(QueryValue("SELECT COUNT(*) FROM [Parametre]"))
    Check "Connexion et requete scalaire", (lngBefore >= 0)

    ' --- Ecritures dans une transaction (annulee a la fin)
    BeginTransaction
    Check "Transaction ouverte", TransactionActive()

    lngId = ExecuteInsert("INSERT INTO [Parametre] ([Cle], [Valeur]) VALUES (?, ?)", "TEST_SQL_1", "valeur 1")
    Check "Insertion : identifiant renvoye", (lngId > 0)

    vVal = QueryValue("SELECT [Valeur] FROM [Parametre] WHERE [ParametreID] = ?", lngId)
    Check "Lecture avec parametre", (TextOf(vVal) = "valeur 1")

    strText = "N'Guessan ""Test"" ; --"
    lngId = ExecuteInsert("INSERT INTO [Parametre] ([Cle], [Valeur]) VALUES (?, ?)", "TEST_SQL_2", strText)
    vVal = QueryValue("SELECT [Valeur] FROM [Parametre] WHERE [ParametreID] = ?", lngId)
    Check "Apostrophe, guillemets et point-virgule", (TextOf(vVal) = strText)

    lngId = ExecuteInsert("INSERT INTO [Parametre] ([Cle], [Valeur]) VALUES (?, ?)", "TEST_SQL_3", Null)
    vVal = QueryValue("SELECT [Valeur] FROM [Parametre] WHERE [ParametreID] = ?", lngId)
    Check "Parametre Null (colonne texte)", IsNull(vVal)

    strText = String$(300, "x")
    lngId = ExecuteInsert("INSERT INTO [Parametre] ([Cle], [Valeur]) VALUES (?, ?)", "TEST_SQL_4", strText)
    vVal = QueryValue("SELECT [Valeur] FROM [Parametre] WHERE [ParametreID] = ?", lngId)
    Check "Texte long (memo de plus de 255 caracteres)", (Len(TextOf(vVal)) = 300)

    ' --- Types Long et Date
    dtTest = DateSerial(2026, 10, 29) + TimeSerial(14, 30, 0)
    lngId = ExecuteInsert("INSERT INTO [VersionSchema] ([Numero], [DateApplication], [Description]) VALUES (?, ?, ?)", _
                          9999, dtTest, "test sql")
    vVal = QueryValue("SELECT [DateApplication] FROM [VersionSchema] WHERE [VersionID] = ?", lngId)
    blnOk = False
    If Not IsNull(vVal) Then blnOk = (CDate(vVal) = dtTest)
    Check "Parametres Long et Date", blnOk

    ' --- Type Oui/Non (Vrai, Faux, Null) : la base impose les liens entre tables,
    '     il faut donc creer d'abord le Role et le Module parents
    lngRole = ExecuteInsert("INSERT INTO [Role] ([Nom]) VALUES (?)", "TEST_SQL_ROLE")
    lngModule = ExecuteInsert("INSERT INTO [Module] ([Nom]) VALUES (?)", "TEST_SQL_MODULE")
    lngId = ExecuteInsert("INSERT INTO [Permission] ([RoleID], [ModuleID], [Lecture], [Ecriture], [Suppression]) " & _
                          "VALUES (?, ?, ?, ?, ?)", lngRole, lngModule, True, False, Null)
    Set AdoRs = QueryRecordset("SELECT [Lecture], [Ecriture], [Suppression] FROM [Permission] WHERE [PermissionID] = ?", lngId)
    Check "Parametre Vrai (colonne Oui/Non)", IsTrueValue(AdoRs.Fields("Lecture").Value)
    Check "Parametre Faux (colonne Oui/Non)", IsFalseValue(AdoRs.Fields("Ecriture").Value)
    Debug.Print "   Null envoye dans une colonne Oui/Non, valeur relue : " & TextOf(AdoRs.Fields("Suppression").Value)
    Check "Null accepte sans erreur dans une colonne Oui/Non (relu : Null ou Faux)", _
        (IsNull(AdoRs.Fields("Suppression").Value) Or IsFalseValue(AdoRs.Fields("Suppression").Value))
    AdoRs.Close

    ' --- Type Currency (la charge exige une categorie et un utilisateur existants)
    lngCategory = ExecuteInsert("INSERT INTO [CategorieCharge] ([Libelle]) VALUES (?)", "TEST_SQL_CATEGORIE")
    lngUser = ExecuteInsert("INSERT INTO [Utilisateur] ([Login], [NomComplet], [MotDePasseHash], [Sel], [RoleID], [Actif], [NbEchecs]) " & _
                            "VALUES (?, ?, ?, ?, ?, ?, ?)", "TEST_SQL_USER", "Test SQL", "hash", "sel", lngRole, True, 0)
    lngId = ExecuteInsert("INSERT INTO [Charge] ([CategorieID], [DateCharge], [Description], [Montant], [UtilisateurID]) " & _
                          "VALUES (?, ?, ?, ?, ?)", lngCategory, dtTest, "test sql", CCur(12345.67), lngUser)
    vVal = QueryValue("SELECT [Montant] FROM [Charge] WHERE [ChargeID] = ?", lngId)
    blnOk = False
    If Not IsNull(vVal) Then blnOk = (CCur(vVal) = CCur(12345.67))
    Check "Parametre Currency", blnOk

    ' --- Recordset deconnecte et mises a jour
    Set AdoRs = QueryRecordset("SELECT [Cle] FROM [Parametre] WHERE [Cle] LIKE ? ORDER BY [Cle]", "TEST_SQL_%")
    Check "Recordset deconnecte : 4 lignes lisibles", ((AdoRs.RecordCount = 4) And (AdoRs.ActiveConnection Is Nothing))
    AdoRs.Close

    lngRows = ExecuteNonQuery("UPDATE [Parametre] SET [Valeur] = ? WHERE [Cle] LIKE ?", "maj", "TEST_SQL_%")
    Check "ExecuteNonQuery : lignes modifiees", (lngRows = 4)

    ' --- Rollback : tout est annule
    RollbackTransaction
    Check "Rollback : transaction fermee", Not TransactionActive()
    Check "Rollback : table Parametre inchangee", (CLng(QueryValue("SELECT COUNT(*) FROM [Parametre]")) = lngBefore)
    Check "Rollback : table VersionSchema inchangee", _
        (CLng(QueryValue("SELECT COUNT(*) FROM [VersionSchema] WHERE [Numero] = ?", 9999)) = 0)

    ' --- Commit : la donnee reste, puis nettoyage
    BeginTransaction
    lngId = ExecuteInsert("INSERT INTO [Parametre] ([Cle], [Valeur]) VALUES (?, ?)", "TEST_SQL_COMMIT", "x")
    CommitTransaction
    Check "Commit : transaction fermee", Not TransactionActive()
    Check "Commit : donnee conservee", _
        (CLng(QueryValue("SELECT COUNT(*) FROM [Parametre] WHERE [Cle] = ?", "TEST_SQL_COMMIT")) = 1)
    Check "Nettoyage : une ligne supprimee", _
        (ExecuteNonQuery("DELETE FROM [Parametre] WHERE [Cle] = ?", "TEST_SQL_COMMIT") = 1)
    Check "Nettoyage : table Parametre revenue a l'etat initial", _
        (CLng(QueryValue("SELECT COUNT(*) FROM [Parametre]")) = lngBefore)

    ' --- Erreurs attendues
    On Error Resume Next

    vVal = QueryValue("SELECT COUNT(*) FROM [Parametre] WHERE [Cle] = ?")
    Check "Refus : nombre de parametres incorrect", (Err.Number = ERR_SQL)
    Err.Clear

    vVal = QueryValue("SELECT COUNT(*) FROM [Parametre] WHERE [Cle] = ?", New Collection)
    Check "Refus : type de parametre non pris en charge", (Err.Number = ERR_SQL)
    Err.Clear

    vVal = QueryValue("SELECT * FROM [TableInexistante]")
    Check "Refus : requete invalide", ((Err.Number = ERR_SQL) And Not TransactionActive())
    Err.Clear

    BeginTransaction
    BeginTransaction
    Check "Refus : deuxieme transaction imbriquee", (Err.Number = ERR_SQL)
    Err.Clear
    RollbackTransaction

    On Error GoTo EH

    Debug.Print "=== " & mlngPassed & " reussis, " & mlngFailed & " echoues ==="
    Exit Sub

EH:
    Debug.Print "ERREUR INATTENDUE (" & Err.Number & ") : " & Err.Description
    mlngFailed = mlngFailed + 1
    RollbackTransaction
    Debug.Print "=== " & mlngPassed & " reussis, " & mlngFailed & " echoues ==="
End Sub
