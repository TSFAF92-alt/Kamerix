Attribute VB_Name = "modKamerixMigration"
Option Explicit
Option Private Module
'==============================================================================
' KAMERIX - Migrations du schema de la base
'------------------------------------------------------------------------------
' MigrateToVersion4 : version 3 -> 4
'     Ajoute la colonne Voiture.MotifRetrait (TEXT 100) : motif saisi quand
'     une voiture est retiree du stock.
' Securites :
'     - une copie du fichier de la base est faite avant toute modification
'       (meme dossier, nom "..._avant_v4_AAAAMMJJ_HHMMSS.accdb") ;
'     - la procedure peut etre relancee sans danger : la colonne n'est ajoutee
'       que si elle n'existe pas, et la version 4 n'est inscrite qu'a la fin ;
'     - la base ne doit pas etre ouverte ailleurs (Access, autre Excel).
' Prerequis : dans modKamerixSchema, SCHEMA_VERSION = 5 et la colonne ajoutee
'     a la table Voiture (pour les bases creees a partir de zero).
' Dependances : modKamerixDb (GetDatabasePath), modKamerixSql.
'==============================================================================

' Erreur de migration (plage 1500 a 1599)
Public Const ERR_MIGRATION As Long = ERR_BASE + 500

Public Sub MigrateToVersion4()
    Dim vVal As Variant
    Dim lngVersion As Long
    Dim strSource As String
    Dim strBackup As String

    On Error GoTo EH

    ' 1. Version actuelle de la base
    vVal = QueryValue("SELECT MAX([Numero]) FROM [VersionSchema]")
    If Not IsNull(vVal) Then lngVersion = CLng(vVal)

    If lngVersion >= 4 Then
        MsgBox "La base est deja en version " & lngVersion & ". Rien a faire.", vbInformation, "Kamerix"
        Exit Sub
    End If
    If lngVersion <> 3 Then
        Err.Raise ERR_MIGRATION, "modKamerixMigration.MigrateToVersion4", _
            "Migration impossible : la base est en version " & lngVersion & " (version 3 attendue)."
    End If

    ' 2. Sauvegarde du fichier avant toute modification
    strSource = GetDatabasePath()
    strBackup = Left$(strSource, InStrRev(strSource, ".") - 1) & "_avant_v4_" & _
                VBA.Format$(VBA.Now, "yyyymmdd_hhnnss") & ".accdb"
    FileCopy strSource, strBackup

    ' 3. Ajout de la colonne (seulement si elle n'existe pas deja)
    If Not ColumnExists("Voiture", "MotifRetrait") Then
        ExecuteNonQuery "ALTER TABLE [Voiture] ADD COLUMN [MotifRetrait] TEXT(100)"
    End If

    ' 4. Inscription de la nouvelle version
    ExecuteInsert "INSERT INTO [VersionSchema] ([Numero], [DateApplication], [Description]) VALUES (?, ?, ?)", _
                  4, Now, "Ajout de Voiture.MotifRetrait"

    MsgBox "Migration terminee : la base est en version 4." & vbCrLf & vbCrLf & _
           "Sauvegarde de l'ancienne version :" & vbCrLf & strBackup, vbInformation, "Kamerix"
    Exit Sub

EH:
    MsgBox "Echec de la migration (erreur " & Err.Number & ") :" & vbCrLf & vbCrLf & Err.Description & _
           vbCrLf & vbCrLf & "Sauvegarde : " & VBA.IIf(Len(strBackup) > 0, strBackup, "(non creee)"), _
           vbCritical, "Kamerix"
End Sub

' Vrai si la colonne existe dans la table (noms internes, jamais saisis par l'utilisateur)
Private Function ColumnExists(ByVal strTable As String, ByVal strColumn As String) As Boolean
    Dim AdoRs As ADODB.Recordset

    On Error Resume Next
    Set AdoRs = QueryRecordset("SELECT TOP 1 [" & strColumn & "] FROM [" & strTable & "]")
    ColumnExists = (Err.Number = 0)
    If Not AdoRs Is Nothing Then AdoRs.Close
    Err.Clear
    On Error GoTo 0
End Function
