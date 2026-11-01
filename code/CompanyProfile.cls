VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CompanyProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Classe CompanyProfile : identité de l'entreprise imprimée sur les documents
' (paramètres SOCIETE_NOM, SOCIETE_ADRESSE, SOCIETE_TELEPHONE, SOCIETE_EMAIL,
' SOCIETE_NIU, SOCIETE_RCCM, SOCIETE_LOGO et PIED_PAGE de la table Parametre).
' Obligatoires : nom, adresse, téléphone. NIU et RCCM sont mis en majuscules.
' LogoPath : chemin tel qu'il est stocké (relatif au dossier de Kamerix, par
' exemple Ressources\logo_20261101_101500.png) ; voir modKamerixCompany.ImportLogo.
' Chargement et enregistrement : modKamerixCompany (LoadCompanyProfile, SaveCompanyProfile).

Private Const MAX_NAME_LENGTH As Long = 100
Private Const MAX_ADDRESS_LENGTH As Long = 200
Private Const MAX_EMAIL_LENGTH As Long = 100
Private Const MAX_LEGAL_ID_LENGTH As Long = 30
Private Const MAX_FOOTER_LENGTH As Long = 150
Private Const MAX_LOGO_PATH_LENGTH As Long = 255

Private mName As String
Private mAddress As String
Private mPhone As String
Private mEmail As String
Private mNiu As String
Private mRccm As String
Private mLogoPath As String
Private mFooter As String

Public Property Get Name() As String
    Name = mName
End Property
Public Property Let Name(ByVal pValue As String)
    mName = Trim$(pValue)
End Property

Public Property Get Address() As String
    Address = mAddress
End Property
Public Property Let Address(ByVal pValue As String)
    mAddress = Trim$(pValue)
End Property

Public Property Get Phone() As String
    Phone = mPhone
End Property
Public Property Let Phone(ByVal pValue As String)
    mPhone = Trim$(pValue)
End Property

Public Property Get Email() As String
    Email = mEmail
End Property
Public Property Let Email(ByVal pValue As String)
    mEmail = Trim$(pValue)
End Property

' Numéro d'identifiant unique (impôts), en majuscules
Public Property Get Niu() As String
    Niu = mNiu
End Property
Public Property Let Niu(ByVal pValue As String)
    mNiu = UCase$(Trim$(pValue))
End Property

' Registre du commerce, en majuscules
Public Property Get Rccm() As String
    Rccm = mRccm
End Property
Public Property Let Rccm(ByVal pValue As String)
    mRccm = UCase$(Trim$(pValue))
End Property

Public Property Get LogoPath() As String
    LogoPath = mLogoPath
End Property
Public Property Let LogoPath(ByVal pValue As String)
    mLogoPath = Trim$(pValue)
End Property

' Mentions libres du pied de facture
Public Property Get Footer() As String
    Footer = mFooter
End Property
Public Property Let Footer(ByVal pValue As String)
    mFooter = Trim$(pValue)
End Property

' ---------- Validation ----------

Public Function ValidationErrors() As Collection
    Dim e As Collection

    Set e = New Collection
    If Len(mName) = 0 Then
        e.Add "Le nom de l'entreprise est obligatoire."
    ElseIf Len(mName) > MAX_NAME_LENGTH Then
        e.Add "Le nom de l'entreprise ne doit pas dépasser " & MAX_NAME_LENGTH & " caractères."
    End If
    If Len(mAddress) = 0 Then
        e.Add "L'adresse est obligatoire."
    ElseIf Len(mAddress) > MAX_ADDRESS_LENGTH Then
        e.Add "L'adresse ne doit pas dépasser " & MAX_ADDRESS_LENGTH & " caractères."
    End If
    If Len(mPhone) = 0 Then
        e.Add "Le téléphone est obligatoire."
    ElseIf Not IsValidPhone(mPhone) Then
        e.Add "Le téléphone est invalide (6 à 15 chiffres ; espaces, +, -, . et parenthèses autorisés)."
    End If
    If Len(mEmail) > 0 Then
        If Len(mEmail) > MAX_EMAIL_LENGTH Or Not IsValidEmail(mEmail) Then e.Add "L'adresse e-mail est invalide."
    End If
    If Len(mNiu) > MAX_LEGAL_ID_LENGTH Then e.Add "Le NIU ne doit pas dépasser " & MAX_LEGAL_ID_LENGTH & " caractères."
    If Len(mRccm) > MAX_LEGAL_ID_LENGTH Then e.Add "Le RCCM ne doit pas dépasser " & MAX_LEGAL_ID_LENGTH & " caractères."
    If Len(mFooter) > MAX_FOOTER_LENGTH Then e.Add "Le pied de page ne doit pas dépasser " & MAX_FOOTER_LENGTH & " caractères."
    If Len(mLogoPath) > MAX_LOGO_PATH_LENGTH Then e.Add "Le chemin du logo est trop long."
    Set ValidationErrors = e
End Function

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
    Err.Raise ERR_COMPANY_INVALID, "CompanyProfile.Validate", msg
End Sub

' E-mail simple : un seul @, du texte avant, un point dans la partie après, aucun espace
Private Function IsValidEmail(ByVal pText As String) As Boolean
    Dim lngAt As Long
    Dim strDomain As String

    If InStr(pText, " ") > 0 Then Exit Function
    lngAt = InStr(pText, "@")
    If lngAt < 2 Then Exit Function
    If InStr(lngAt + 1, pText, "@") > 0 Then Exit Function
    strDomain = Mid$(pText, lngAt + 1)
    IsValidEmail = (InStr(strDomain, ".") > 1) And (Right$(strDomain, 1) <> ".")
End Function
