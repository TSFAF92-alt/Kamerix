Attribute VB_Name = "modKamerixDb"
Option Explicit
Option Private Module
'==============================================================================
' KAMERIX - Creation et ouverture de la base de donnees Kamerixdb.accdb
'------------------------------------------------------------------------------
' Role        : creer (une seule fois) la base Access protegee par mot de passe,
'               avec tout son schema, puis fournir une connexion ADO au reste
'               de l'application.
' Emplacement : <dossier de ThisWorkbook>\Database\Kamerixdb.accdb
'               (le dossier Database est cree automatiquement)
' Format      : ACCDB Access 2007 via le fournisseur Microsoft.ACE.OLEDB.12.0.
'               Le mot de passe est pose a la creation (chiffrement "ancien
'               mode"), celui que le code ADO sait ouvrir. Ne jamais proteger
'               la base depuis Access 2010+ avec le chiffrement par defaut.
' Liaison     : ADODB en liaison precoce (reference "Microsoft ActiveX Data
'               Objects" cochee dans le projet). ADOX.Catalog et
'               Scripting.FileSystemObject restent en liaison tardive.
' Requetes    : les requetes generiques (lecture, ecriture, transactions)
'               sont dans modKamerixSql.
' Atomicite   : la base est construite dans un fichier temporaire, verifiee,
'               puis renommee. Un echec ne laisse aucun fichier a moitie fait.
' Dependance  : modKamerixSchema (DDL, donnees de reference, liste des tables).
' Erreurs     : toutes les procedures remontent l'erreur a l'appelant avec
'               Err.Raise (numeros ERR_* ci-dessous). Aucune MsgBox ici, sauf
'               dans la procedure de test TestDatabaseCreation.
'==============================================================================

'--- Parametres de la base ----------------------------------------------------
Private Const PROVIDER As String = "Microsoft.ACE.OLEDB.12.0"
Private Const DB_FOLDER As String = "Database"
Private Const DB_FILENAME As String = "Kamerixdb.accdb"
Private Const DB_TEMP_FILENAME As String = "Kamerixdb_creation.accdb"

'>>> A PERSONNALISER AVANT LA PREMIERE CREATION (1 a 20 caracteres, sans ; ' ") <<<
Private Const DB_PASSWORD As String = "@KamerixAdmin2026"
Private Const PASSWORD_PLACEHOLDER As String = "CHANGER_CE_MOT_DE_PASSE"

'--- Numeros d'erreur personnalises -------------------------------------------
Public Const ERR_BASE As Long = vbObjectError + 1000
Public Const ERR_NO_PATH As Long = ERR_BASE + 1
Public Const ERR_PASSWORD As Long = ERR_BASE + 2
Public Const ERR_SQL As Long = ERR_BASE + 3
Public Const ERR_VERIFY As Long = ERR_BASE + 4

'==============================================================================
' API PUBLIQUE
'==============================================================================

' Dossier de la base : <dossier du classeur>\Database
Public Function GetDatabaseFolder() As String
    Dim strBase As String

    strBase = ThisWorkbook.Path
    If Len(strBase) = 0 Then
        Err.Raise ERR_NO_PATH, "modKamerixDb.GetDatabaseFolder", _
            "Le classeur n'est pas enregistre sur le disque : le dossier de la base ne peut pas etre determine."
    End If
    If LCase$(Left$(strBase, 4)) = "http" Then
        Err.Raise ERR_NO_PATH, "modKamerixDb.GetDatabaseFolder", _
            "Le classeur est dans un dossier en ligne (" & strBase & "). Access ne peut pas ouvrir une base a cette adresse : deplacez le classeur dans un dossier local."
    End If
    If Right$(strBase, 1) = "\" Then strBase = Left$(strBase, Len(strBase) - 1)

    GetDatabaseFolder = strBase & "\" & DB_FOLDER
End Function

' Chemin complet de la base
Public Function GetDatabasePath() As String
    GetDatabasePath = GetDatabaseFolder() & "\" & DB_FILENAME
End Function

' Vrai si le fichier de la base existe deja
Public Function DatabaseExists() As Boolean
    Dim fso As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    DatabaseExists = fso.FileExists(GetDatabasePath())
End Function

' Garantit que la base existe et est valide.
'   - Absente  : elle est creee, remplie (schema + donnees de reference), verifiee.
'                Retourne True.
'   - Presente : elle est ouverte et controlee (mot de passe, tables, version).
'                Elle n'est JAMAIS modifiee ni ecrasee. Retourne False.
Public Function EnsureDatabase() As Boolean
    Dim fso As Object
    Dim strFolder As String
    Dim strFinal As String
    Dim strTemp As String
    Dim blnTempStarted As Boolean
    Dim lngErr As Long
    Dim strErr As String
    Dim strSrc As String

    On Error GoTo EH

    Set fso = CreateObject("Scripting.FileSystemObject")
    strFolder = GetDatabaseFolder()
    strFinal = GetDatabasePath()

    If Not fso.FolderExists(strFolder) Then fso.CreateFolder strFolder

    If fso.FileExists(strFinal) Then
        VerifyDatabase strFinal, False
        EnsureDatabase = False
        Exit Function
    End If

    ValidatePassword

    strTemp = strFolder & "\" & DB_TEMP_FILENAME
    If fso.FileExists(strTemp) Then fso.DeleteFile strTemp, True

    blnTempStarted = True
    CreateDatabaseFile strTemp
    BuildSchema strTemp
    VerifyDatabase strTemp, True

    fso.MoveFile strTemp, strFinal
    blnTempStarted = False

    EnsureDatabase = True
    Exit Function

EH:
    lngErr = Err.Number
    strErr = Err.Description
    strSrc = Err.Source

    On Error Resume Next
    If blnTempStarted Then
        If fso.FileExists(strTemp) Then fso.DeleteFile strTemp, True
    End If
    On Error GoTo 0

    Err.Raise lngErr, strSrc, strErr
End Function

' Ouvre et retourne une connexion ADO sur la base. L'appelant doit la fermer :
'     Adocn.Close : Set Adocn = Nothing
' Pour lire ou ecrire des donnees, preferer les fonctions de modKamerixSql
' (QueryRecordset, QueryValue, ExecuteNonQuery, ExecuteInsert) qui gerent
' la connexion et les parametres.
Public Function OpenConnection(Optional ByVal blnReadOnly As Boolean = False) As ADODB.Connection
    Dim Adocn As ADODB.Connection
    Dim lngErr As Long
    Dim strErr As String
    Dim strSrc As String

    On Error GoTo EH

    Set Adocn = New ADODB.Connection
    Adocn.Open BuildConnectionString(GetDatabasePath(), blnReadOnly)
    Set OpenConnection = Adocn
    Exit Function

EH:
    lngErr = Err.Number
    strErr = Err.Description & vbCrLf & ConnectionHint()
    strSrc = Err.Source

    On Error Resume Next
    CloseQuietly Adocn
    On Error GoTo 0

    Err.Raise lngErr, strSrc, strErr
End Function

'==============================================================================
' PROCEDURE DE TEST (a lancer avec F5 depuis l'editeur VBA)
'==============================================================================
Public Sub TestDatabaseCreation()
    Dim blnCreated As Boolean

    On Error GoTo EH

    blnCreated = EnsureDatabase()
    If blnCreated Then
        MsgBox "Base creee et verifiee :" & vbCrLf & GetDatabasePath(), vbInformation, "Kamerix"
    Else
        MsgBox "La base existe deja et est valide :" & vbCrLf & GetDatabasePath(), vbInformation, "Kamerix"
    End If
    Exit Sub

EH:
    MsgBox "Echec (erreur " & Err.Number & ") :" & vbCrLf & vbCrLf & Err.Description, vbCritical, "Kamerix"
End Sub

'==============================================================================
' PROCEDURES INTERNES
'==============================================================================

' Chaine de connexion (mot de passe inclus)
Private Function BuildConnectionString(ByVal strPath As String, Optional ByVal blnReadOnly As Boolean = False) As String
    Dim strCn As String

    strCn = "Provider=" & PROVIDER & ";Data Source=" & strPath & _
            ";Jet OLEDB:Database Password=" & DB_PASSWORD & ";"
    If blnReadOnly Then strCn = strCn & "Mode=Read;"

    BuildConnectionString = strCn
End Function

Private Function ConnectionHint() As String
    ConnectionHint = "Verifier : (1) le mot de passe de la base, " & _
        "(2) que le moteur Microsoft Access Database Engine (ACE 12.0) est installe, " & _
        "(3) que sa version 32 ou 64 bits correspond a celle d'Excel."
End Function

Private Sub ValidatePassword()
    Dim strMsg As String

    If DB_PASSWORD = PASSWORD_PLACEHOLDER Then
        strMsg = "Definissez le mot de passe de la base (constante DB_PASSWORD dans modKamerixDb) avant la creation."
    ElseIf Len(DB_PASSWORD) > 20 Then
        strMsg = "Le mot de passe d'une base Access (ancien mode de chiffrement) est limite a 20 caracteres."
    ElseIf InStr(DB_PASSWORD, ";") > 0 Or InStr(DB_PASSWORD, "'") > 0 Or InStr(DB_PASSWORD, """") > 0 Then
        strMsg = "Le mot de passe ne doit contenir ni point-virgule, ni apostrophe, ni guillemet."
    End If

    If Len(strMsg) > 0 Then Err.Raise ERR_PASSWORD, "modKamerixDb.ValidatePassword", strMsg
End Sub

' Cree le fichier .accdb vide, protege par mot de passe
Private Sub CreateDatabaseFile(ByVal strPath As String)
    Dim cat As Object
    Dim lngErr As Long
    Dim strErr As String
    Dim strSrc As String

    On Error GoTo EH

    Set cat = CreateObject("ADOX.Catalog")
    cat.Create "Provider=" & PROVIDER & ";Data Source=" & strPath & _
               ";Jet OLEDB:Database Password=" & DB_PASSWORD & ";"

    ' Catalog.Create laisse une connexion ouverte : on la ferme
    cat.ActiveConnection.Close
    Set cat = Nothing
    Exit Sub

EH:
    lngErr = Err.Number
    strErr = Err.Description & vbCrLf & ConnectionHint()
    strSrc = Err.Source

    On Error Resume Next
    If Not cat Is Nothing Then cat.ActiveConnection.Close
    Set cat = Nothing
    On Error GoTo 0

    Err.Raise lngErr, strSrc, strErr
End Sub

' Ouvre la base temporaire et applique le schema complet
Private Sub BuildSchema(ByVal strPath As String)
    Dim Adocn As ADODB.Connection
    Dim lngErr As Long
    Dim strErr As String
    Dim strSrc As String

    On Error GoTo EH

    Set Adocn = New ADODB.Connection
    Adocn.Open BuildConnectionString(strPath)
    ApplySchema Adocn
    Adocn.Close
    Set Adocn = Nothing
    Exit Sub

EH:
    lngErr = Err.Number
    strErr = Err.Description
    strSrc = Err.Source

    On Error Resume Next
    CloseQuietly Adocn
    On Error GoTo 0

    Err.Raise lngErr, strSrc, strErr
End Sub

' Controle une base : ouverture avec mot de passe, presence de toutes les tables,
' version du schema. Si blnCheckProtection : verifie aussi qu'elle refuse
' l'ouverture SANS mot de passe (a faire a la creation).
Private Sub VerifyDatabase(ByVal strPath As String, ByVal blnCheckProtection As Boolean)
    Dim Adocn As ADODB.Connection
    Dim cat As Object
    Dim AdoRs As ADODB.Recordset
    Dim vNames As Variant
    Dim i As Long
    Dim strMissing As String
    Dim lngVersion As Long
    Dim lngErr As Long
    Dim strErr As String
    Dim strSrc As String

    On Error GoTo EH

    Set Adocn = New ADODB.Connection
    Adocn.Open BuildConnectionString(strPath, True)

    Set cat = CreateObject("ADOX.Catalog")
    Set cat.ActiveConnection = Adocn

    vNames = ExpectedTables()
    For i = LBound(vNames) To UBound(vNames)
        If Not TableExists(cat, CStr(vNames(i))) Then
            strMissing = strMissing & vNames(i) & ", "
        End If
    Next i
    If Len(strMissing) > 0 Then
        Err.Raise ERR_VERIFY, "modKamerixDb.VerifyDatabase", _
            "Tables manquantes dans la base : " & Left$(strMissing, Len(strMissing) - 2)
    End If

    Set AdoRs = Adocn.Execute("SELECT MAX([Numero]) FROM [VersionSchema]")
    If Not AdoRs.EOF Then
        If Not IsNull(AdoRs.Fields(0).Value) Then lngVersion = CLng(AdoRs.Fields(0).Value)
    End If
    AdoRs.Close
    Set AdoRs = Nothing

    If lngVersion <> SCHEMA_VERSION Then
        Err.Raise ERR_VERIFY, "modKamerixDb.VerifyDatabase", _
            "Version du schema incorrecte : base en version " & lngVersion & _
            ", code en version " & SCHEMA_VERSION & ". Une migration est necessaire."
    End If

    Set cat = Nothing
    Adocn.Close
    Set Adocn = Nothing

    If blnCheckProtection Then
        If OpensWithoutPassword(strPath) Then
            Err.Raise ERR_VERIFY, "modKamerixDb.VerifyDatabase", _
                "La base s'ouvre sans mot de passe : la protection n'a pas ete appliquee."
        End If
    End If
    Exit Sub

EH:
    lngErr = Err.Number
    strErr = Err.Description
    strSrc = Err.Source

    On Error Resume Next
    If Not AdoRs Is Nothing Then AdoRs.Close
    Set cat = Nothing
    CloseQuietly Adocn
    On Error GoTo 0

    Err.Raise lngErr, strSrc, strErr
End Sub

Private Function TableExists(ByVal cat As Object, ByVal strName As String) As Boolean
    Dim tbl As Object

    For Each tbl In cat.Tables
        If StrComp(tbl.Name, strName, vbTextCompare) = 0 Then
            TableExists = True
            Exit Function
        End If
    Next tbl
End Function

' Vrai si la base s'ouvre SANS fournir de mot de passe (ce qui serait anormal)
Private Function OpensWithoutPassword(ByVal strPath As String) As Boolean
    Dim Adocn As ADODB.Connection

    On Error Resume Next
    Set Adocn = New ADODB.Connection
    Adocn.Open "Provider=" & PROVIDER & ";Data Source=" & strPath & ";Mode=Read;"
    OpensWithoutPassword = (Err.Number = 0)
    Err.Clear
    CloseQuietly Adocn
    Set Adocn = Nothing
    On Error GoTo 0
End Function

' Ferme une connexion sans jamais lever d'erreur
Private Sub CloseQuietly(ByVal Adocn As ADODB.Connection)
    On Error Resume Next
    If Not Adocn Is Nothing Then
        If Adocn.State <> adStateClosed Then Adocn.Close
    End If
    On Error GoTo 0
End Sub
