VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Party"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Classe Party (tiers) : acheteur et/ou vendeur (table Tiers).
' Les règles de validation sont ici, pas dans les formulaires.
' La vérification du rôle selon le type de mouvement (acheteur pour une vente, vendeur pour un achat)
' sera faite par la couche qui crée le mouvement.

Private Const MAX_NAME_LENGTH As Long = 120
Private Const MAX_PHONE_LENGTH As Long = 30
Private Const MAX_ID_TYPE_LENGTH As Long = 30
Private Const MAX_ID_NUMBER_LENGTH As Long = 40
Private Const MAX_ADDRESS_LENGTH As Long = 200

Private mIdentifier As Long
Private mName As String
Private mPhone As String
Private mIdType As String
Private mIdNumber As String
Private mAddress As String
Private mIsBuyer As Boolean
Private mIsSeller As Boolean

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

Public Property Get IdType() As String
    IdType = mIdType
End Property
Public Property Let IdType(ByVal pValue As String)
    mIdType = Trim$(pValue)
End Property

Public Property Get IdNumber() As String
    IdNumber = mIdNumber
End Property
Public Property Let IdNumber(ByVal pValue As String)
    mIdNumber = Trim$(pValue)
End Property

Public Property Get Address() As String
    Address = mAddress
End Property
Public Property Let Address(ByVal pValue As String)
    mAddress = Trim$(pValue)
End Property

Public Property Get IsBuyer() As Boolean
    IsBuyer = mIsBuyer
End Property
Public Property Let IsBuyer(ByVal pValue As Boolean)
    mIsBuyer = pValue
End Property

Public Property Get IsSeller() As Boolean
    IsSeller = mIsSeller
End Property
Public Property Let IsSeller(ByVal pValue As Boolean)
    mIsSeller = pValue
End Property

' Charge l'état lu dans la base (les colonnes vides arrivent en Null)
Friend Sub Hydrate(ByVal pId As Long, ByVal pName As Variant, ByVal pPhone As Variant, _
                    ByVal pIdType As Variant, ByVal pIdNumber As Variant, _
                    ByVal pAddress As Variant, ByVal pIsBuyer As Variant, _
                    ByVal pIsSeller As Variant)
    mIdentifier = pId
    mName = TextOf(pName)
    mPhone = TextOf(pPhone)
    mIdType = TextOf(pIdType)
    mIdNumber = TextOf(pIdNumber)
    mAddress = TextOf(pAddress)
    mIsBuyer = BoolOf(pIsBuyer)
    mIsSeller = BoolOf(pIsSeller)
End Sub

' Appelé par la classe d'accès aux données après l'insertion
Friend Sub SetId(ByVal pId As Long)
    mIdentifier = pId
End Sub

' Liste de tous les problèmes trouvés (collection vide = tiers valide)
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

    If (Len(mIdType) > 0) <> (Len(mIdNumber) > 0) Then
        e.Add "Le type et le numéro de la pièce d'identité vont ensemble : renseigner les deux ou aucun."
    End If
    If Len(mIdType) > MAX_ID_TYPE_LENGTH Then
        e.Add "Le type de pièce ne doit pas dépasser " & MAX_ID_TYPE_LENGTH & " caractères."
    End If
    If Len(mIdNumber) > MAX_ID_NUMBER_LENGTH Then
        e.Add "Le numéro de pièce ne doit pas dépasser " & MAX_ID_NUMBER_LENGTH & " caractères."
    End If

    If Len(mAddress) > MAX_ADDRESS_LENGTH Then
        e.Add "L'adresse ne doit pas dépasser " & MAX_ADDRESS_LENGTH & " caractères."
    End If

    If Not (mIsBuyer Or mIsSeller) Then
        e.Add "Un tiers doit être acheteur, vendeur, ou les deux."
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
    Err.Raise ERR_PARTY_INVALID, "Party.Validate", msg
End Sub

' ---------- Méthodes privées ----------
' (le contrôle du téléphone est partagé : modKamerixValidation.IsValidPhone)


Private Function TextOf(ByVal pValue As Variant) As String
    If IsNull(pValue) Or IsEmpty(pValue) Then
        TextOf = ""
    Else
        TextOf = Trim$(CStr(pValue))
    End If
End Function

Private Function BoolOf(ByVal pValue As Variant) As Boolean
    If IsNull(pValue) Or IsEmpty(pValue) Then
        BoolOf = False
    Else
        BoolOf = CBool(pValue)
    End If
End Function
