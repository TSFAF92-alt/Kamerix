Attribute VB_Name = "modKamerixReference"
Option Explicit
Option Private Module

'==============================================================================
' KAMERIX - Referentiel : marques, modeles, lieux, liste des intermediaires
'------------------------------------------------------------------------------
' Role        : ajouter, renommer, supprimer et lister les donnees de reference
'               utilisees par les listes deroulantes (pas de faute de saisie
'               dans les statistiques).
' Droits      : non controles ici (couche donnees) ; l'ecran ou le service qui
'               appelle verifie le droit d'ecriture (module Voitures ou Tiers).
' Suppression : refusee (ERR_REF_IN_USE) tant que l'element est utilise ;
'               la base le refuserait de toute facon (relations sans cascade).
' Lieux       : VILLE (sans parent), QUARTIER (parent VILLE obligatoire),
'               COMMISSARIAT (parent VILLE facultatif).
' Erreurs     : ERR_REF_INVALID, ERR_REF_DUPLICATE, ERR_REF_NOT_FOUND, ERR_REF_IN_USE.
' Dependances : modKamerixSql, modKamerixConst.
'==============================================================================

Private Const MAX_BRAND_LENGTH As Long = 60
Private Const MAX_MODEL_LENGTH As Long = 60
Private Const MAX_PLACE_LENGTH As Long = 100

'==============================================================================
' MARQUES
'==============================================================================

Public Function AddBrand(ByVal pName As String) As Long
    Dim strName As String

    strName = CleanName(pName, MAX_BRAND_LENGTH, "Le nom de la marque", "modKamerixReference.AddBrand")
    If CountOf("SELECT COUNT(*) FROM [Marque] WHERE [Nom] = ?", strName) > 0 Then
        RaiseDuplicate "Cette marque existe d" & ChrW(233) & "j" & ChrW(224) & " : " & strName, "modKamerixReference.AddBrand"
    End If
    AddBrand = ExecuteInsert("INSERT INTO [Marque] ([Nom]) VALUES (?)", strName)
End Function

Public Sub RenameBrand(ByVal pBrandId As Long, ByVal pName As String)
    Dim strName As String

    strName = CleanName(pName, MAX_BRAND_LENGTH, "Le nom de la marque", "modKamerixReference.RenameBrand")
    If CountOf("SELECT COUNT(*) FROM [Marque] WHERE [Nom] = ? AND [MarqueID] <> ?", strName, pBrandId) > 0 Then
        RaiseDuplicate "Cette marque existe d" & ChrW(233) & "j" & ChrW(224) & " : " & strName, "modKamerixReference.RenameBrand"
    End If
    If ExecuteNonQuery("UPDATE [Marque] SET [Nom] = ? WHERE [MarqueID] = ?", strName, pBrandId) = 0 Then
        RaiseNotFound "Marque", pBrandId, "modKamerixReference.RenameBrand"
    End If
End Sub

Public Sub DeleteBrand(ByVal pBrandId As Long)
    If CountOf("SELECT COUNT(*) FROM [Modele] WHERE [MarqueID] = ?", pBrandId) > 0 Then
        RaiseInUse "Cette marque a des mod" & ChrW(232) & "les : supprimez-les d'abord.", "modKamerixReference.DeleteBrand"
    End If
    If ExecuteNonQuery("DELETE FROM [Marque] WHERE [MarqueID] = ?", pBrandId) = 0 Then
        RaiseNotFound "Marque", pBrandId, "modKamerixReference.DeleteBrand"
    End If
End Sub

' Colonnes : MarqueID, Nom (tri par nom)
Public Function BrandList() As ADODB.Recordset
    Set BrandList = QueryRecordset("SELECT [MarqueID], [Nom] FROM [Marque] ORDER BY [Nom]")
End Function

'==============================================================================
' MODELES
'==============================================================================

Public Function AddModel(ByVal pBrandId As Long, ByVal pName As String) As Long
    Dim strName As String

    strName = CleanName(pName, MAX_MODEL_LENGTH, "Le nom du mod" & ChrW(232) & "le", "modKamerixReference.AddModel")
    If CountOf("SELECT COUNT(*) FROM [Marque] WHERE [MarqueID] = ?", pBrandId) = 0 Then
        RaiseNotFound "Marque", pBrandId, "modKamerixReference.AddModel"
    End If
    If CountOf("SELECT COUNT(*) FROM [Modele] WHERE [MarqueID] = ? AND [Nom] = ?", pBrandId, strName) > 0 Then
        RaiseDuplicate "Ce mod" & ChrW(232) & "le existe d" & ChrW(233) & "j" & ChrW(224) & " pour cette marque : " & strName, _
                       "modKamerixReference.AddModel"
    End If
    AddModel = ExecuteInsert("INSERT INTO [Modele] ([MarqueID], [Nom]) VALUES (?, ?)", pBrandId, strName)
End Function

Public Sub RenameModel(ByVal pModelId As Long, ByVal pName As String)
    Dim strName As String
    Dim vBrand As Variant

    strName = CleanName(pName, MAX_MODEL_LENGTH, "Le nom du mod" & ChrW(232) & "le", "modKamerixReference.RenameModel")
    vBrand = QueryValue("SELECT [MarqueID] FROM [Modele] WHERE [ModeleID] = ?", pModelId)
    If IsNull(vBrand) Then RaiseNotFound "Mod" & ChrW(232) & "le", pModelId, "modKamerixReference.RenameModel"
    If CountOf("SELECT COUNT(*) FROM [Modele] WHERE [MarqueID] = ? AND [Nom] = ? AND [ModeleID] <> ?", _
               CLng(vBrand), strName, pModelId) > 0 Then
        RaiseDuplicate "Ce mod" & ChrW(232) & "le existe d" & ChrW(233) & "j" & ChrW(224) & " pour cette marque : " & strName, _
                       "modKamerixReference.RenameModel"
    End If
    ExecuteNonQuery "UPDATE [Modele] SET [Nom] = ? WHERE [ModeleID] = ?", strName, pModelId
End Sub

Public Sub DeleteModel(ByVal pModelId As Long)
    If CountOf("SELECT COUNT(*) FROM [Voiture] WHERE [ModeleID] = ?", pModelId) > 0 Then
        RaiseInUse "Ce mod" & ChrW(232) & "le est utilis" & ChrW(233) & " par des voitures.", "modKamerixReference.DeleteModel"
    End If
    If ExecuteNonQuery("DELETE FROM [Modele] WHERE [ModeleID] = ?", pModelId) = 0 Then
        RaiseNotFound "Mod" & ChrW(232) & "le", pModelId, "modKamerixReference.DeleteModel"
    End If
End Sub

' Colonnes : ModeleID, MarqueID, NomMarque, Nom. pBrandId = 0 : tous les modeles.
Public Function ModelList(Optional ByVal pBrandId As Long = 0) As ADODB.Recordset
    Const SQL_BASE As String = "SELECT m.[ModeleID], m.[MarqueID], b.[Nom] AS NomMarque, m.[Nom] " & _
                               "FROM [Modele] AS m INNER JOIN [Marque] AS b ON m.[MarqueID] = b.[MarqueID] "
    If pBrandId = 0 Then
        Set ModelList = QueryRecordset(SQL_BASE & "ORDER BY b.[Nom], m.[Nom]")
    Else
        Set ModelList = QueryRecordset(SQL_BASE & "WHERE m.[MarqueID] = ? ORDER BY m.[Nom]", pBrandId)
    End If
End Function

' Marque d'un modele (0 si le modele n'existe pas)
Public Function BrandOfModel(ByVal pModelId As Long) As Long
    Dim vBrand As Variant

    vBrand = QueryValue("SELECT [MarqueID] FROM [Modele] WHERE [ModeleID] = ?", pModelId)
    If Not IsNull(vBrand) Then BrandOfModel = CLng(vBrand)
End Function

' Libelle complet d'un modele : "Toyota Corolla" (chaine vide si introuvable)
Public Function ModelLabel(ByVal pModelId As Long) As String
    Dim vLabel As Variant

    vLabel = QueryValue("SELECT b.[Nom] & ' ' & m.[Nom] FROM [Modele] AS m " & _
                        "INNER JOIN [Marque] AS b ON m.[MarqueID] = b.[MarqueID] WHERE m.[ModeleID] = ?", pModelId)
    If Not IsNull(vLabel) Then ModelLabel = CStr(vLabel)
End Function

'==============================================================================
' LIEUX
'==============================================================================

' pParentId = 0 : pas de lieu parent
Public Function AddPlace(ByVal pName As String, ByVal pPlaceType As String, Optional ByVal pParentId As Long = 0) As Long
    Dim strName As String
    Dim vParent As Variant

    strName = CleanName(pName, MAX_PLACE_LENGTH, "Le nom du lieu", "modKamerixReference.AddPlace")
    CheckPlaceRules pPlaceType, pParentId

    If pParentId = 0 Then
        vParent = Null
        If CountOf("SELECT COUNT(*) FROM [Lieu] WHERE [Nom] = ? AND [TypeLieu] = ? AND [LieuParentID] IS NULL", _
                   strName, pPlaceType) > 0 Then
            RaiseDuplicate "Ce lieu existe d" & ChrW(233) & "j" & ChrW(224) & " : " & strName, "modKamerixReference.AddPlace"
        End If
    Else
        vParent = pParentId
        If CountOf("SELECT COUNT(*) FROM [Lieu] WHERE [Nom] = ? AND [TypeLieu] = ? AND [LieuParentID] = ?", _
                   strName, pPlaceType, pParentId) > 0 Then
            RaiseDuplicate "Ce lieu existe d" & ChrW(233) & "j" & ChrW(224) & " : " & strName, "modKamerixReference.AddPlace"
        End If
    End If

    AddPlace = ExecuteInsert("INSERT INTO [Lieu] ([Nom], [TypeLieu], [LieuParentID]) VALUES (?, ?, ?)", _
                             strName, pPlaceType, vParent)
End Function

Public Sub RenamePlace(ByVal pPlaceId As Long, ByVal pName As String)
    Dim strName As String

    strName = CleanName(pName, MAX_PLACE_LENGTH, "Le nom du lieu", "modKamerixReference.RenamePlace")
    If ExecuteNonQuery("UPDATE [Lieu] SET [Nom] = ? WHERE [LieuID] = ?", strName, pPlaceId) = 0 Then
        RaiseNotFound "Lieu", pPlaceId, "modKamerixReference.RenamePlace"
    End If
End Sub

Public Sub DeletePlace(ByVal pPlaceId As Long)
    If CountOf("SELECT COUNT(*) FROM [Lieu] WHERE [LieuParentID] = ?", pPlaceId) > 0 _
       Or CountOf("SELECT COUNT(*) FROM [Intermediaire] WHERE [VilleID] = ?", pPlaceId) > 0 _
       Or CountOf("SELECT COUNT(*) FROM [Mouvement] WHERE [LieuID] = ?", pPlaceId) > 0 _
       Or CountOf("SELECT COUNT(*) FROM [SuiviCV] WHERE [CommissariatID] = ?", pPlaceId) > 0 Then
        RaiseInUse "Ce lieu est utilis" & ChrW(233) & " (sous-lieux, interm" & ChrW(233) & "diaires, mouvements ou CV).", _
                   "modKamerixReference.DeletePlace"
    End If
    If ExecuteNonQuery("DELETE FROM [Lieu] WHERE [LieuID] = ?", pPlaceId) = 0 Then
        RaiseNotFound "Lieu", pPlaceId, "modKamerixReference.DeletePlace"
    End If
End Sub

' Colonnes : LieuID, Nom, TypeLieu, LieuParentID, NomParent.
' Filtres facultatifs : type de lieu, lieu parent (0 = tous).
Public Function PlaceList(Optional ByVal pPlaceType As String = "", Optional ByVal pParentId As Long = 0) As ADODB.Recordset
    Const SQL_BASE As String = "SELECT l.[LieuID], l.[Nom], l.[TypeLieu], l.[LieuParentID], p.[Nom] AS NomParent " & _
                               "FROM [Lieu] AS l LEFT JOIN [Lieu] AS p ON l.[LieuParentID] = p.[LieuID] "
    If Len(pPlaceType) > 0 And pParentId <> 0 Then
        Set PlaceList = QueryRecordset(SQL_BASE & "WHERE l.[TypeLieu] = ? AND l.[LieuParentID] = ? ORDER BY l.[Nom]", _
                                       pPlaceType, pParentId)
    ElseIf Len(pPlaceType) > 0 Then
        Set PlaceList = QueryRecordset(SQL_BASE & "WHERE l.[TypeLieu] = ? ORDER BY l.[Nom]", pPlaceType)
    ElseIf pParentId <> 0 Then
        Set PlaceList = QueryRecordset(SQL_BASE & "WHERE l.[LieuParentID] = ? ORDER BY l.[Nom]", pParentId)
    Else
        Set PlaceList = QueryRecordset(SQL_BASE & "ORDER BY l.[TypeLieu], l.[Nom]")
    End If
End Function

' Type d'un lieu (chaine vide s'il n'existe pas)
Public Function PlaceTypeOf(ByVal pPlaceId As Long) As String
    Dim vType As Variant

    vType = QueryValue("SELECT [TypeLieu] FROM [Lieu] WHERE [LieuID] = ?", pPlaceId)
    If Not IsNull(vType) Then PlaceTypeOf = CStr(vType)
End Function

'==============================================================================
' INTERMEDIAIRES (liste ; lecture et enregistrement : classe DataAccess)
'==============================================================================

' Colonnes : IntermediaireID, Nom, Telephone, VilleID, NomVille (tri par nom).
' pSearch : partie du nom (vide = tous)
Public Function IntermediaryList(Optional ByVal pSearch As String = "") As ADODB.Recordset
    Set IntermediaryList = QueryRecordset( _
        "SELECT i.[IntermediaireID], i.[Nom], i.[Telephone], i.[VilleID], l.[Nom] AS NomVille " & _
        "FROM [Intermediaire] AS i INNER JOIN [Lieu] AS l ON i.[VilleID] = l.[LieuID] " & _
        "WHERE i.[Nom] LIKE ? ORDER BY i.[Nom]", "%" & Trim$(pSearch) & "%")
End Function

'==============================================================================
' PROCEDURES INTERNES
'==============================================================================

' Regles de type et de parent d'un lieu
Private Sub CheckPlaceRules(ByVal pPlaceType As String, ByVal pParentId As Long)
    Const SRC As String = "modKamerixReference.AddPlace"
    Dim strParentType As String

    Select Case pPlaceType
        Case PLACE_TYPE_CITY, PLACE_TYPE_DISTRICT, PLACE_TYPE_POLICE_STATION
        Case Else
            Err.Raise ERR_REF_INVALID, SRC, "Type de lieu inconnu : " & pPlaceType
    End Select

    If pParentId <> 0 Then
        strParentType = PlaceTypeOf(pParentId)
        If Len(strParentType) = 0 Then RaiseNotFound "Lieu parent", pParentId, SRC
    End If

    Select Case pPlaceType
        Case PLACE_TYPE_CITY
            If pParentId <> 0 Then Err.Raise ERR_REF_INVALID, SRC, "Une ville n'a pas de lieu parent."
        Case PLACE_TYPE_DISTRICT
            If strParentType <> PLACE_TYPE_CITY Then
                Err.Raise ERR_REF_INVALID, SRC, "Un quartier doit " & ChrW(234) & "tre rattach" & ChrW(233) & " " & ChrW(224) & " une ville."
            End If
        Case PLACE_TYPE_POLICE_STATION
            If pParentId <> 0 And strParentType <> PLACE_TYPE_CITY Then
                Err.Raise ERR_REF_INVALID, SRC, "Un commissariat ne peut " & ChrW(234) & "tre rattach" & ChrW(233) & _
                    " qu'" & ChrW(224) & " une ville."
            End If
    End Select
End Sub

' Nom nettoye et controle (obligatoire, longueur maximale)
Private Function CleanName(ByVal pName As String, ByVal pMaxLength As Long, ByVal pLabel As String, _
                           ByVal pSource As String) As String
    Dim strName As String

    strName = Trim$(pName)
    If Len(strName) = 0 Then Err.Raise ERR_REF_INVALID, pSource, pLabel & " est obligatoire."
    If Len(strName) > pMaxLength Then
        Err.Raise ERR_REF_INVALID, pSource, pLabel & " ne doit pas d" & ChrW(233) & "passer " & pMaxLength & " caract" & ChrW(232) & "res."
    End If
    CleanName = strName
End Function

Private Function CountOf(ByVal pSql As String, ParamArray pParams() As Variant) As Long
    CountOf = CLng(RunSql(pSql, pParams, kxScalar))
End Function

Private Sub RaiseDuplicate(ByVal pMessage As String, ByVal pSource As String)
    Err.Raise ERR_REF_DUPLICATE, pSource, pMessage
End Sub

Private Sub RaiseNotFound(ByVal pWhat As String, ByVal pId As Long, ByVal pSource As String)
    Err.Raise ERR_REF_NOT_FOUND, pSource, pWhat & " introuvable (identifiant " & pId & ")."
End Sub

Private Sub RaiseInUse(ByVal pMessage As String, ByVal pSource As String)
    Err.Raise ERR_REF_IN_USE, pSource, "Suppression impossible : " & pMessage
End Sub
