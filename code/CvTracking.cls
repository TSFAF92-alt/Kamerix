VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CvTracking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Classe CvTracking : suivi du Certificat de Vente d'une vente (table SuiviCV).
' Le CV est établi au commissariat ; Kamerix n'en assure que le suivi.
' Statut : A_DEPOSER (à la création de la vente) -> EN_COURS (dépôt, vente
' soldée exigée : contrôle du service) -> SIGNE (la voiture passe à « vendu »).
' Les frais payés sont un coût interne de la vente (jamais facturés au client).

Private Const MAX_CV_NUMBER_LENGTH As Long = 40

Private mIdentifier As Long
Private mMovementId As Long
Private mCvNumber As String
Private mPoliceStationId As Long
Private mFilingDate As Variant          ' Null tant que le dossier n'est pas déposé
Private mSignatureDate As Variant       ' Null tant que le CV n'est pas signé
Private mFeesPaid As Currency
Private mStatus As String

Private Sub Class_Initialize()
    mStatus = CV_STATUS_TO_FILE
    mFilingDate = Null
    mSignatureDate = Null
End Sub

Public Property Get Id() As Long
    Id = mIdentifier
End Property
Public Property Get MovementId() As Long
    MovementId = mMovementId
End Property
Public Property Let MovementId(ByVal pValue As Long)
    mMovementId = pValue
End Property
Public Property Get CvNumber() As String
    CvNumber = mCvNumber
End Property
Public Property Get PoliceStationId() As Long
    PoliceStationId = mPoliceStationId
End Property
Public Property Get FilingDate() As Variant
    FilingDate = mFilingDate
End Property
Public Property Get SignatureDate() As Variant
    SignatureDate = mSignatureDate
End Property
Public Property Get FeesPaid() As Currency
    FeesPaid = mFeesPaid
End Property
Public Property Get Status() As String
    Status = mStatus
End Property

' A_DEPOSER -> EN_COURS : dossier déposé au commissariat, frais payés
Public Sub Submit(ByVal pPoliceStationId As Long, ByVal pFilingDate As Date, ByVal pFees As Currency)
    If mStatus <> CV_STATUS_TO_FILE Then
        Err.Raise ERR_CV, "CvTracking.Submit", "Le dossier de CV a déjà été déposé (statut : " & mStatus & ")."
    End If
    If pPoliceStationId <= 0 Then Err.Raise ERR_CV, "CvTracking.Submit", "Le commissariat est obligatoire."
    If pFees < 0 Then Err.Raise ERR_CV, "CvTracking.Submit", "Les frais du CV ne peuvent pas être négatifs."
    If pFilingDate > Date + 1 Then Err.Raise ERR_CV, "CvTracking.Submit", "La date de dépôt ne peut pas être dans le futur."
    mPoliceStationId = pPoliceStationId
    mFilingDate = pFilingDate
    mFeesPaid = pFees
    mStatus = CV_STATUS_IN_PROGRESS
End Sub

' EN_COURS -> SIGNE : CV signé, avec son numéro
Public Sub MarkSigned(ByVal pCvNumber As String, ByVal pSignatureDate As Date)
    If mStatus <> CV_STATUS_IN_PROGRESS Then
        Err.Raise ERR_CV, "CvTracking.MarkSigned", "Seul un dossier de CV en cours peut être signé (statut : " & mStatus & ")."
    End If
    If Len(Trim$(pCvNumber)) = 0 Then Err.Raise ERR_CV, "CvTracking.MarkSigned", "Le numéro du CV est obligatoire."
    If Len(Trim$(pCvNumber)) > MAX_CV_NUMBER_LENGTH Then
        Err.Raise ERR_CV, "CvTracking.MarkSigned", "Le numéro du CV ne doit pas dépasser " & MAX_CV_NUMBER_LENGTH & " caractères."
    End If
    If pSignatureDate < CDate(mFilingDate) Then
        Err.Raise ERR_CV, "CvTracking.MarkSigned", "La date de signature ne peut pas précéder la date de dépôt."
    End If
    If pSignatureDate > Date + 1 Then Err.Raise ERR_CV, "CvTracking.MarkSigned", "La date de signature ne peut pas être dans le futur."
    mCvNumber = Trim$(pCvNumber)
    mSignatureDate = pSignatureDate
    mStatus = CV_STATUS_SIGNED
End Sub

Friend Sub Hydrate(ByVal pId As Long, ByVal pMovementId As Long, ByVal pCvNumber As Variant, ByVal pPoliceStationId As Variant, _
                   ByVal pFilingDate As Variant, ByVal pSignatureDate As Variant, ByVal pFeesPaid As Currency, _
                   ByVal pStatus As String)
    mIdentifier = pId
    mMovementId = pMovementId
    mCvNumber = TextOfValue(pCvNumber)
    If IsNull(pPoliceStationId) Then mPoliceStationId = 0 Else mPoliceStationId = CLng(pPoliceStationId)
    mFilingDate = pFilingDate
    mSignatureDate = pSignatureDate
    mFeesPaid = pFeesPaid
    mStatus = pStatus
End Sub

Friend Sub SetId(ByVal pId As Long)
    mIdentifier = pId
End Sub
