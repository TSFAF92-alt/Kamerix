Attribute VB_Name = "modKamerixConst"
Option Explicit
Option Private Module
' Codes de statut de la voiture : identiques à ceux stockés dans la base
Public Const CAR_STATUS_IN_STOCK As String = "EN_STOCK"
Public Const CAR_STATUS_RESERVED As String = "RESERVE"
Public Const CAR_STATUS_SALE_IN_PROGRESS As String = "VENTE_EN_COURS"
Public Const CAR_STATUS_SOLD As String = "VENDU"
Public Const CAR_STATUS_WITHDRAWN As String = "RETIRE"

' Numéros d'erreur métier (1100 à 1199 : classe Car)
Public Const ERR_CAR_TRANSITION As Long = vbObjectError + 1100
Public Const ERR_CAR_REASON As Long = vbObjectError + 1101
Public Const ERR_CAR_STATUS As Long = vbObjectError + 1102
Public Const ERR_CAR_DATA As Long = vbObjectError + 1103
Public Const ERR_CAR_LOCKED As Long = vbObjectError + 1104

' Numéros d'erreur métier (1200 à 1299 : personnes, classes Party et Intermediary)
Public Const ERR_PARTY_INVALID As Long = vbObjectError + 1200
Public Const ERR_INTERMEDIARY_INVALID As Long = vbObjectError + 1210

' Numéros d'erreur métier (1300 à 1399 : accès aux données)
Public Const ERR_DATA_DUPLICATE As Long = vbObjectError + 1300
Public Const ERR_DATA_NOT_FOUND As Long = vbObjectError + 1301

' Numéros d'erreur métier (1600 à 1699 : référentiel et paramètres)
Public Const ERR_REF_INVALID As Long = vbObjectError + 1600      ' donnée invalide
Public Const ERR_REF_DUPLICATE As Long = vbObjectError + 1601    ' nom déjà utilisé
Public Const ERR_REF_NOT_FOUND As Long = vbObjectError + 1602    ' élément introuvable
Public Const ERR_REF_IN_USE As Long = vbObjectError + 1603       ' suppression impossible : élément utilisé
Public Const ERR_SETTING As Long = vbObjectError + 1650          ' paramètre absent ou invalide

' Numéros d'erreur métier (1700 à 1799 : numérotation des pièces)
Public Const ERR_NUMBERING As Long = vbObjectError + 1700

' Numéros d'erreur métier (1800 à 1899 : mouvements, achat et vente)
Public Const ERR_MOVEMENT_INVALID As Long = vbObjectError + 1800  ' données du mouvement invalides
Public Const ERR_MOVEMENT_STATUS As Long = vbObjectError + 1801   ' changement de statut interdit
Public Const ERR_PURCHASE As Long = vbObjectError + 1810          ' règle de l'achat non respectée
Public Const ERR_SALE As Long = vbObjectError + 1820              ' règle de la vente non respectée
Public Const ERR_PAYMENT_INVALID As Long = vbObjectError + 1830   ' données du paiement invalides
Public Const ERR_INVOICE_INVALID As Long = vbObjectError + 1840   ' données de la facture invalides
Public Const ERR_CV As Long = vbObjectError + 1850                ' suivi du certificat de vente

' Mouvement : type, mode de règlement, statut (valeurs stockées dans la base)
Public Const MOVEMENT_TYPE_PURCHASE As String = "ACHAT"
Public Const MOVEMENT_TYPE_SALE As String = "VENTE"
Public Const SETTLEMENT_CASH As String = "COMPTANT"
Public Const SETTLEMENT_INSTALLMENT As String = "ECHELONNE"
Public Const MOVEMENT_STATUS_IN_PROGRESS As String = "EN_COURS"
Public Const MOVEMENT_STATUS_SETTLED As String = "SOLDE"
Public Const MOVEMENT_STATUS_CANCELLED As String = "ANNULE"

' Paiement : sens
Public Const PAYMENT_DIRECTION_IN As String = "ENCAISSEMENT"
Public Const PAYMENT_DIRECTION_REFUND As String = "REMBOURSEMENT"

' Facture : statut
Public Const INVOICE_STATUS_ISSUED As String = "EMISE"
Public Const INVOICE_STATUS_CANCELLED As String = "ANNULEE"

' Suivi du certificat de vente : statut
Public Const CV_STATUS_TO_FILE As String = "A_DEPOSER"
Public Const CV_STATUS_IN_PROGRESS As String = "EN_COURS"
Public Const CV_STATUS_SIGNED As String = "SIGNE"

' Types de pièces numérotées (table Compteur) : préfixe du numéro FAC-2026-00001
Public Const DOC_INVOICE As String = "FAC"
Public Const DOC_RECEIPT As String = "REC"
Public Const DOC_REFUND As String = "RMB"

' Types de lieux (table Lieu)
Public Const PLACE_TYPE_CITY As String = "VILLE"
Public Const PLACE_TYPE_DISTRICT As String = "QUARTIER"
Public Const PLACE_TYPE_POLICE_STATION As String = "COMMISSARIAT"
