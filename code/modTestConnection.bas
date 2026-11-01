Attribute VB_Name = "modTestConnection"
Option Explicit
Option Private Module
'==============================================================================
' Test de connexion uniquement (aucune ecriture dans la base)
' Lancer TestConnection avec F5, puis lire la fenetre Execution.
' Utilise modKamerixDb : GetDatabasePath, DatabaseExists, OpenConnection.
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

Public Sub TestConnection()
    Dim Adocn As ADODB.Connection
    Dim AdoRs As ADODB.Recordset
    Dim lngTables As Long

    mlngPassed = 0
    mlngFailed = 0
    Debug.Print "=== Test de connexion a la base ==="

    ' 1. Le chemin et le fichier
    On Error GoTo EH_PATH
    Debug.Print "Chemin : " & GetDatabasePath()
    Check "Le fichier de la base existe", DatabaseExists()
    If Not DatabaseExists() Then GoTo FINISH

    ' 2. Ouverture en lecture/ecriture
    On Error GoTo EH_OPEN
    Set Adocn = OpenConnection()
    Check "Ouverture en lecture/ecriture (mot de passe accepte)", (Adocn.State = adStateOpen)

    ' 3. Une requete simple sur la table des versions
    On Error GoTo EH_QUERY
    Set AdoRs = Adocn.Execute("SELECT MAX([Numero]) FROM [VersionSchema]")
    Check "Lecture de VersionSchema", (Not AdoRs.EOF)
    If Not AdoRs.EOF Then
        If Not IsNull(AdoRs.Fields(0).Value) Then Debug.Print "Version du schema : " & AdoRs.Fields(0).Value
    End If
    AdoRs.Close
    Set AdoRs = Nothing

    ' 4. Fermeture propre
    Adocn.Close
    Check "Fermeture de la connexion", (Adocn.State = adStateClosed)
    Set Adocn = Nothing

    ' 5. Ouverture en lecture seule
    On Error GoTo EH_OPEN
    Set Adocn = OpenConnection(True)
    Check "Ouverture en lecture seule", (Adocn.State = adStateOpen)
    Adocn.Close
    Set Adocn = Nothing

FINISH:
    Debug.Print "=== " & mlngPassed & " reussis, " & mlngFailed & " echoues ==="
    Exit Sub

EH_PATH:
    Debug.Print "ECHEC - chemin de la base : " & Err.Description
    mlngFailed = mlngFailed + 1
    Resume FINISH

EH_OPEN:
    Debug.Print "ECHEC - ouverture de la base (erreur " & Err.Number & ") : " & Err.Description
    mlngFailed = mlngFailed + 1
    Resume FINISH

EH_QUERY:
    Debug.Print "ECHEC - requete (erreur " & Err.Number & ") : " & Err.Description
    mlngFailed = mlngFailed + 1
    On Error Resume Next
    If Not Adocn Is Nothing Then Adocn.Close
    On Error GoTo 0
    Resume FINISH
End Sub
