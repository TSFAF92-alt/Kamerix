Attribute VB_Name = "modKamerixParties"
Option Explicit
Option Private Module

'==============================================================================
' KAMERIX - Service des tiers et des intermediaires (ecrans du lot 5.4b)
'------------------------------------------------------------------------------
' Listes      : PartyList (recherche par nom, filtre par role), lue par les
'               ecrans et par les listes deroulantes (acheteurs, vendeurs).
' Ecriture    : SaveParty / SaveIntermediary via DataAccess, apres controle du
'               droit d'ecriture sur le module Tiers.
' SQL         : seulement des fonctions reconnues par le moteur ACE via ADO (IIf, IsNull,
'               Trim) ; Nz n'existe que dans l'application Access.
' Dependances : DataAccess, Party, Intermediary, modKamerixSession, modKamerixSql.
'==============================================================================

Public Const PARTY_ROLE_ALL As String = ""
Public Const PARTY_ROLE_BUYER As String = "ACHETEUR"
Public Const PARTY_ROLE_SELLER As String = "VENDEUR"

' Liste des tiers. Colonnes : TiersID, Nom, Telephone, Piece, Roles.
'   pSearch : partie du nom (vide = tous) ; pRole : PARTY_ROLE_* (vide = tous)
' Lecture : module Tiers.
Public Function PartyList(Optional ByVal pSearch As String = "", Optional ByVal pRole As String = PARTY_ROLE_ALL) As ADODB.Recordset
    Dim strSql As String
    Dim strRoleFilter As String

    RequireRead MOD_PARTIES

    Select Case pRole
        Case PARTY_ROLE_BUYER:  strRoleFilter = " AND [EstAcheteur] = True"
        Case PARTY_ROLE_SELLER: strRoleFilter = " AND [EstVendeur] = True"
    End Select

    strSql = "SELECT [TiersID], [Nom], [Telephone], " & _
             "Trim(IIf(IsNull([TypePiece]), '', [TypePiece]) & ' ' & IIf(IsNull([NumeroPiece]), '', [NumeroPiece])) AS Piece, " & _
             "IIf([EstAcheteur] And [EstVendeur], 'Acheteur et vendeur', IIf([EstAcheteur], 'Acheteur', 'Vendeur')) AS Roles " & _
             "FROM [Tiers] WHERE [Nom] LIKE ?" & strRoleFilter & " ORDER BY [Nom]"
    Set PartyList = QueryRecordset(strSql, "%" & Trim$(pSearch) & "%")
End Function

' Enregistre un tiers (nouveau ou existant). Ecriture : module Tiers.
Public Sub SavePartyChecked(ByVal pParty As Party)
    Dim objData As DataAccess

    RequireWrite MOD_PARTIES
    Set objData = New DataAccess
    objData.SaveParty pParty
End Sub

' Tiers par identifiant (Nothing s'il n'existe pas). Lecture : module Tiers.
Public Function LoadPartyChecked(ByVal pPartyId As Long) As Party
    Dim objData As DataAccess

    RequireRead MOD_PARTIES
    Set objData = New DataAccess
    Set LoadPartyChecked = objData.LoadParty(pPartyId)
End Function

' Enregistre un intermediaire (nouveau ou existant). Ecriture : module Tiers.
Public Sub SaveIntermediaryChecked(ByVal pIntermediary As Intermediary)
    Dim objData As DataAccess

    RequireWrite MOD_PARTIES
    Set objData = New DataAccess
    objData.SaveIntermediary pIntermediary
End Sub

' Intermediaire par identifiant (Nothing s'il n'existe pas). Lecture : module Tiers.
Public Function LoadIntermediaryChecked(ByVal pIntermediaryId As Long) As Intermediary
    Dim objData As DataAccess

    RequireRead MOD_PARTIES
    Set objData = New DataAccess
    Set LoadIntermediaryChecked = objData.LoadIntermediary(pIntermediaryId)
End Function
