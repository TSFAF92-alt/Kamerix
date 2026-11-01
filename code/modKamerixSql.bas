Attribute VB_Name = "modKamerixSql"
Option Explicit
Option Private Module
'==============================================================================
' KAMERIX - Acces generique aux donnees (ADO, liaison precoce)
'------------------------------------------------------------------------------
' Role        : point d'entree unique pour toutes les requetes SQL de
'               l'application. Gere la connexion, les parametres et les
'               transactions. Aucune autre classe ne construit de SQL "a la
'               main" avec des valeurs collees dans le texte.
' Parametres  : le SQL utilise des "?" et les valeurs sont passees a la suite.
'               Plus d'apostrophes a doubler, aucun risque d'injection SQL.
'                 QueryValue "SELECT [Nom] FROM [Tiers] WHERE [TiersID] = ?", lngId
' Connexion   : ouverte puis fermee a chaque appel (modKamerixDb.OpenConnection).
'               Pendant une transaction, la connexion de la transaction est
'               reutilisee pour tous les appels.
' Types       : String, Long/Integer/Byte, Single/Double, Currency, Boolean,
'               Date, Null/Empty (Null SQL). Autre type : erreur ERR_SQL.
' Erreurs     : toutes les erreurs de requete remontent avec le numero
'               ERR_SQL (modKamerixDb) et le texte de la requete.
' Dependance  : modKamerixDb (OpenConnection, ERR_SQL).
'==============================================================================

' Mode d'execution de la fonction generique RunSql
Public Enum KxSqlMode
    kxRecordset = 1     ' SELECT : retourne un Recordset deconnecte
    kxNonQuery = 2      ' INSERT/UPDATE/DELETE : retourne le nombre de lignes touchees
    kxScalar = 3        ' SELECT d'une valeur : retourne la 1re valeur (Null si aucune ligne)
    kxInsert = 4        ' INSERT : retourne l'identifiant NumeroAuto genere
End Enum

' Connexion de la transaction en cours (Nothing si aucune)
Private mcnTrans As ADODB.Connection

'==============================================================================
' FONCTION GENERIQUE
'==============================================================================

' Execute une requete SQL parametree selon le mode demande.
'   strSql  : requete avec des "?" pour les valeurs
'   vParams : tableau des valeurs, dans l'ordre des "?" (Array() si aucune)
'   eMode   : voir KxSqlMode
Public Function RunSql(ByVal strSql As String, ByVal vParams As Variant, ByVal eMode As KxSqlMode) As Variant
    Dim Adocn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim AdoRs As ADODB.Recordset
    Dim blnOwnsConnection As Boolean
    Dim lngAffected As Long
    Dim lngErr As Long
    Dim strErr As String
    Dim strSrc As String

    ' Connexion : celle de la transaction en cours, sinon une connexion temporaire.
    ' (hors du gestionnaire d'erreur : une erreur de connexion garde son numero)
    If Not mcnTrans Is Nothing Then
        Set Adocn = mcnTrans
    Else
        Set Adocn = OpenConnection()
        blnOwnsConnection = True
    End If

    On Error GoTo EH

    Set cmd = BuildCommand(Adocn, strSql, vParams)

    Select Case eMode
        Case kxRecordset
            Set AdoRs = New ADODB.Recordset
            AdoRs.CursorLocation = adUseClient
            AdoRs.Open cmd, , adOpenStatic, adLockReadOnly
            Set AdoRs.ActiveConnection = Nothing      ' deconnecte : utilisable apres fermeture de Adocn
            Set RunSql = AdoRs

        Case kxNonQuery
            cmd.Execute lngAffected, , adExecuteNoRecords
            RunSql = lngAffected

        Case kxScalar
            Set AdoRs = cmd.Execute
            If AdoRs.EOF Then
                RunSql = Null
            Else
                RunSql = AdoRs.Fields(0).Value
            End If
            AdoRs.Close

        Case kxInsert
            cmd.Execute , , adExecuteNoRecords
            Set AdoRs = Adocn.Execute("SELECT @@IDENTITY")
            RunSql = CLng(AdoRs.Fields(0).Value)
            AdoRs.Close

        Case Else
            Err.Raise ERR_SQL, "modKamerixSql.RunSql", "Mode d'execution inconnu : " & eMode
    End Select

    If blnOwnsConnection Then CloseConnectionQuietly Adocn
    Exit Function

EH:
    lngErr = Err.Number
    strErr = Err.Description
    strSrc = Err.Source

    On Error Resume Next
    If eMode <> kxRecordset Then
        If Not AdoRs Is Nothing Then AdoRs.Close
    End If
    If blnOwnsConnection Then CloseConnectionQuietly Adocn
    On Error GoTo 0

    If lngErr = ERR_SQL Then Err.Raise ERR_SQL, strSrc, strErr
    Err.Raise ERR_SQL, "modKamerixSql.RunSql", _
        "Erreur SQL (" & lngErr & ") : " & strErr & vbCrLf & "Requete : " & strSql
End Function

'==============================================================================
' RACCOURCIS (a utiliser dans le reste de l'application)
'==============================================================================

' SELECT : Recordset deconnecte (lecture seule). L'appelant le ferme : AdoRs.Close
Public Function QueryRecordset(ByVal strSql As String, ParamArray vParams() As Variant) As ADODB.Recordset
    Set QueryRecordset = RunSql(strSql, vParams, kxRecordset)
End Function

' SELECT d'une valeur : retourne Null s'il n'y a aucune ligne ou si la valeur est Null
Public Function QueryValue(ByVal strSql As String, ParamArray vParams() As Variant) As Variant
    QueryValue = RunSql(strSql, vParams, kxScalar)
End Function

' INSERT / UPDATE / DELETE : retourne le nombre de lignes touchees
Public Function ExecuteNonQuery(ByVal strSql As String, ParamArray vParams() As Variant) As Long
    ExecuteNonQuery = CLng(RunSql(strSql, vParams, kxNonQuery))
End Function

' INSERT : retourne l'identifiant NumeroAuto de la ligne creee
Public Function ExecuteInsert(ByVal strSql As String, ParamArray vParams() As Variant) As Long
    ExecuteInsert = CLng(RunSql(strSql, vParams, kxInsert))
End Function

' Chaine vide ou espaces -> Null (pour ne pas stocker de texte vide dans les colonnes facultatives)
Public Function NullIfEmpty(ByVal vValue As Variant) As Variant
    If IsNull(vValue) Or IsEmpty(vValue) Then
        NullIfEmpty = Null
    ElseIf VarType(vValue) = vbString Then
        If Len(Trim$(vValue)) = 0 Then
            NullIfEmpty = Null
        Else
            NullIfEmpty = vValue
        End If
    Else
        NullIfEmpty = vValue
    End If
End Function

'==============================================================================
' TRANSACTIONS
'------------------------------------------------------------------------------
' Modele d'utilisation (tout ou rien) :
'     BeginTransaction
'     On Error GoTo EH
'     ... ExecuteInsert / ExecuteNonQuery ...
'     CommitTransaction
'     Exit Sub
' EH: RollbackTransaction          ' sans effet s'il n'y a pas de transaction
'     Err.Raise ...
'==============================================================================

Public Sub BeginTransaction()
    If Not mcnTrans Is Nothing Then
        Err.Raise ERR_SQL, "modKamerixSql.BeginTransaction", "Une transaction est deja en cours."
    End If

    Set mcnTrans = OpenConnection()
    mcnTrans.BeginTrans
End Sub

Public Sub CommitTransaction()
    Dim lngErr As Long
    Dim strErr As String
    Dim strSrc As String

    If mcnTrans Is Nothing Then
        Err.Raise ERR_SQL, "modKamerixSql.CommitTransaction", "Aucune transaction en cours."
    End If

    On Error GoTo EH
    mcnTrans.CommitTrans
    CloseConnectionQuietly mcnTrans
    Set mcnTrans = Nothing
    Exit Sub

EH:
    lngErr = Err.Number
    strErr = Err.Description
    strSrc = Err.Source

    On Error Resume Next
    mcnTrans.RollbackTrans
    CloseConnectionQuietly mcnTrans
    Set mcnTrans = Nothing
    On Error GoTo 0

    Err.Raise lngErr, strSrc, strErr
End Sub

' Annule la transaction en cours. Sans effet (et sans erreur) s'il n'y en a pas.
Public Sub RollbackTransaction()
    If mcnTrans Is Nothing Then Exit Sub

    On Error Resume Next
    mcnTrans.RollbackTrans
    CloseConnectionQuietly mcnTrans
    Set mcnTrans = Nothing
    On Error GoTo 0
End Sub

Public Function TransactionActive() As Boolean
    TransactionActive = Not (mcnTrans Is Nothing)
End Function

'==============================================================================
' PROCEDURES INTERNES
'==============================================================================

' Prepare la commande : controle le nombre de parametres puis les ajoute
Private Function BuildCommand(ByVal Adocn As ADODB.Connection, ByVal strSql As String, ByVal vParams As Variant) As ADODB.Command
    Dim cmd As ADODB.Command
    Dim lngPlaceholders As Long
    Dim lngCount As Long
    Dim i As Long

    lngPlaceholders = Len(strSql) - Len(Replace(strSql, "?", ""))
    If IsArray(vParams) Then lngCount = UBound(vParams) - LBound(vParams) + 1

    If lngPlaceholders <> lngCount Then
        Err.Raise ERR_SQL, "modKamerixSql.BuildCommand", _
            "Nombre de parametres incorrect : " & lngPlaceholders & " attendu(s), " & lngCount & " fourni(s)." & _
            vbCrLf & "Requete : " & strSql
    End If

    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = Adocn
    cmd.CommandText = strSql
    cmd.CommandType = adCmdText

    If lngCount > 0 Then
        For i = LBound(vParams) To UBound(vParams)
            cmd.Parameters.Append CreateParam(cmd, vParams(i))
        Next i
    End If

    Set BuildCommand = cmd
End Function

' Cree le parametre ADO adapte au type de la valeur VBA
Private Function CreateParam(ByVal cmd As ADODB.Command, ByVal vValue As Variant) As ADODB.Parameter
    Dim lngLen As Long

    Select Case VarType(vValue)
        Case vbNull, vbEmpty
            Set CreateParam = cmd.CreateParameter(, adVarWChar, adParamInput, 1, Null)

        Case vbString
            lngLen = Len(vValue)
            If lngLen > 255 Then
                Set CreateParam = cmd.CreateParameter(, adLongVarWChar, adParamInput, lngLen, vValue)
            Else
                If lngLen = 0 Then lngLen = 1
                Set CreateParam = cmd.CreateParameter(, adVarWChar, adParamInput, lngLen, vValue)
            End If

        Case vbByte, vbInteger, vbLong
            Set CreateParam = cmd.CreateParameter(, adInteger, adParamInput, , CLng(vValue))

        Case vbSingle, vbDouble, vbDecimal
            Set CreateParam = cmd.CreateParameter(, adDouble, adParamInput, , CDbl(vValue))

        Case vbCurrency
            Set CreateParam = cmd.CreateParameter(, adCurrency, adParamInput, , vValue)

        Case vbBoolean
            Set CreateParam = cmd.CreateParameter(, adBoolean, adParamInput, , vValue)

        Case vbDate
            Set CreateParam = cmd.CreateParameter(, adDate, adParamInput, , vValue)

        Case Else
            Err.Raise ERR_SQL, "modKamerixSql.CreateParam", _
                "Type de parametre non pris en charge (VarType = " & VarType(vValue) & ")."
    End Select
End Function

' Ferme une connexion sans jamais lever d'erreur
Private Sub CloseConnectionQuietly(ByVal Adocn As ADODB.Connection)
    On Error Resume Next
    If Not Adocn Is Nothing Then
        If Adocn.State <> adStateClosed Then Adocn.Close
    End If
    On Error GoTo 0
End Sub
