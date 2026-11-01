Attribute VB_Name = "modKamerixCars"
Option Explicit
Option Private Module

'==============================================================================
' KAMERIX - Service des voitures (ecrans du lot 5.4b)
'------------------------------------------------------------------------------
' Liste       : CarList, filtree par statut, avec le libelle du statut et le
'               prix d'achat (achat non annule).
' Fiche       : SaveCarDetails (donnees descriptives ; une voiture vendue ou
'               retiree n'est plus modifiable : controle de la classe Car).
' Retrait     : WithdrawCar (en stock -> retiree, motif obligatoire).
' Droits      : lecture ou ecriture sur le module Voitures.
' Dependances : DataAccess, Car, modKamerixSession, modKamerixSql.
'==============================================================================

' Liste des voitures. Colonnes : VoitureID, Modele, Annee, Chassis, Immatriculation,
' Statut (code), LibelleStatut, PrixAchat. pStatus : code de statut, vide = toutes.
' Lecture : module Voitures.
Public Function CarList(Optional ByVal pStatus As String = "") As ADODB.Recordset
    Dim strSql As String

    RequireRead MOD_CARS

    strSql = "SELECT v.[VoitureID], b.[Nom] & ' ' & m.[Nom] AS Modele, v.[Annee], v.[Chassis], v.[Immatriculation], " & _
             "v.[Statut], " & StatusLabelSql() & " AS LibelleStatut, " & _
             "(SELECT TOP 1 a.[Prix] FROM [Mouvement] AS a WHERE a.[VoitureID] = v.[VoitureID] " & _
             "AND a.[TypeMouvement] = ? AND a.[Statut] <> ?) AS PrixAchat " & _
             "FROM ([Voiture] AS v INNER JOIN [Modele] AS m ON v.[ModeleID] = m.[ModeleID]) " & _
             "INNER JOIN [Marque] AS b ON m.[MarqueID] = b.[MarqueID] "

    If Len(pStatus) = 0 Then
        Set CarList = QueryRecordset(strSql & "ORDER BY b.[Nom], m.[Nom], v.[Chassis]", _
                                     MOVEMENT_TYPE_PURCHASE, MOVEMENT_STATUS_CANCELLED)
    Else
        Set CarList = QueryRecordset(strSql & "WHERE v.[Statut] = ? ORDER BY b.[Nom], m.[Nom], v.[Chassis]", _
                                     MOVEMENT_TYPE_PURCHASE, MOVEMENT_STATUS_CANCELLED, pStatus)
    End If
End Function

' Libelle affiche d'un statut de voiture
Public Function CarStatusLabel(ByVal pStatus As String) As String
    Select Case pStatus
        Case CAR_STATUS_IN_STOCK:         CarStatusLabel = "En stock"
        Case CAR_STATUS_RESERVED:         CarStatusLabel = "R" & ChrW(233) & "serv" & ChrW(233) & "e"
        Case CAR_STATUS_SALE_IN_PROGRESS: CarStatusLabel = "Vente en cours"
        Case CAR_STATUS_SOLD:             CarStatusLabel = "Vendue"
        Case CAR_STATUS_WITHDRAWN:        CarStatusLabel = "Retir" & ChrW(233) & "e"
        Case Else:                        CarStatusLabel = pStatus
    End Select
End Function

' Voiture par identifiant (Nothing si elle n'existe pas). Lecture : module Voitures.
Public Function LoadCarChecked(ByVal pCarId As Long) As Car
    Dim objData As DataAccess

    RequireRead MOD_CARS
    Set objData = New DataAccess
    Set LoadCarChecked = objData.LoadCar(pCarId)
End Function

' Enregistre les donnees descriptives d'une voiture existante. Ecriture : module Voitures.
' (Une nouvelle voiture s'enregistre avec son achat : modKamerixPurchase.RecordPurchase.)
Public Sub SaveCarDetails(ByVal pCar As Car)
    Dim objData As DataAccess

    RequireWrite MOD_CARS
    If pCar.Id = 0 Then
        Err.Raise ERR_PURCHASE, "modKamerixCars.SaveCarDetails", _
            "Une nouvelle voiture s'enregistre avec son achat."
    End If
    Set objData = New DataAccess
    objData.SaveCar pCar
End Sub

' Retire une voiture du stock (motif obligatoire, 100 caracteres au plus). Ecriture : module Voitures.
Public Sub WithdrawCar(ByVal pCarId As Long, ByVal pReason As String)
    Dim objData As DataAccess
    Dim objCar As Car

    RequireWrite MOD_CARS
    Set objData = New DataAccess
    Set objCar = objData.LoadCar(pCarId)
    If objCar Is Nothing Then
        Err.Raise ERR_DATA_NOT_FOUND, "modKamerixCars.WithdrawCar", "Voiture introuvable (identifiant " & pCarId & ")."
    End If
    objCar.Withdraw pReason
    objData.SaveCar objCar
End Sub

' Expression SQL du libelle de statut (memes libelles que CarStatusLabel)
Private Function StatusLabelSql() As String
    StatusLabelSql = "Switch(v.[Statut] = '" & CAR_STATUS_IN_STOCK & "', 'En stock', " & _
                     "v.[Statut] = '" & CAR_STATUS_RESERVED & "', '" & CarStatusLabel(CAR_STATUS_RESERVED) & "', " & _
                     "v.[Statut] = '" & CAR_STATUS_SALE_IN_PROGRESS & "', 'Vente en cours', " & _
                     "v.[Statut] = '" & CAR_STATUS_SOLD & "', 'Vendue', " & _
                     "v.[Statut] = '" & CAR_STATUS_WITHDRAWN & "', '" & CarStatusLabel(CAR_STATUS_WITHDRAWN) & "')"
End Function
