VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Intermediary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Classe Intermediary (intermédiaire, démarcheur) : table Intermediaire.
' Il est rattaché à une ville ; sa commission est portée par le mouvement
' (achat ou vente), pas par l'intermédiaire.

Private Const MAX_NAME_LENGTH As Long = 120
Private Const MAX_PHONE_LENGTH As Long = 30

Private mIdentifier As Long
Private mName As String
Private mPhone As String
Private mCityId As Long

Public Property Get Id() As Long
    Id = mIdentifier
End Property

Public Property Get Name() As String
    Name = mName
End Property
Public Property Let Name(ByVal pValue As String)
    mName = Trim$(pValue)
End Property

Public Property Get Phone() As String
    Phone = mPhone
End Property
Public Property Let Phone(ByVal pValue As String)
    mPhone = Trim$(pValue)
End Property

' Ville de l'intermédiaire (lieu de type VILLE)
Public Property Get CityId() As Long
    CityId = mCityId
End Property
Public Property Let CityId(ByVal pValue As Long)
    mCityId = pValue
End Property

' Charge l'état lu dans la base (classe d'accès aux données uniquement)
Friend Sub Hydrate(ByVal pId As Long, ByVal pName As Variant, ByVal pPhone As Variant, ByVal pCityId As Long)
    mIdentifier = pId
    mName = TextOfValue(pName)
    mPhone = TextOfValue(pPhone)
    mCityId = pCityId
End Sub

' Appelé par la classe d'accès aux données après l'insertion
Friend Sub SetId(ByVal pId As Long)
    mIdentifier = pId
End Sub

' Liste de tous les problèmes trouvés (collection vide = intermédiaire valide)
Public Function ValidationErrors() As Collection
    Dim e As Collection

    Set e = New Collection

    If Len(mName) = 0 Then
        e.Add "Le nom est obligatoire."
    ElseIf Len(mName) > MAX_NAME_LENGTH Then
        e.Add "Le nom ne doit pas dépasser " & MAX_NAME_LENGTH & " caractères."
    End If

    If Len(mPhone) > MAX_PHONE_LENGTH Then
        e.Add "Le téléphone ne doit pas dépasser " & MAX_PHONE_LENGTH & " caractères."
    ElseIf Len(mPhone) > 0 Then
        If Not IsValidPhone(mPhone) Then
            e.Add "Le téléphone doit contenir de 6 à 15 chiffres (espaces, +, -, . et parenthèses autorisés)."
        End If
    End If

    If mCityId <= 0 Then
        e.Add "La ville est obligatoire."
    End If

    Set ValidationErrors = e
End Function

' Lève une erreur unique regroupant tous les problèmes (à appeler avant l'enregistrement)
Public Sub Validate()
    Dim e As Collection
    Dim msg As String
    Dim i As Long

    Set e = ValidationErrors()
    If e.Count = 0 Then Exit Sub

    For i = 1 To e.Count
        If i > 1 Then msg = msg & vbCrLf
        msg = msg & "- " & e.Item(i)
    Next i
    Err.Raise ERR_INTERMEDIARY_INVALID, "Intermediary.Validate", msg
End Sub
