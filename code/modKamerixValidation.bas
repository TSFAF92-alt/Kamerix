Attribute VB_Name = "modKamerixValidation"
Option Explicit
Option Private Module

'==============================================================================
' KAMERIX - Controles de saisie partages par plusieurs classes
'------------------------------------------------------------------------------
' Fonctions pures (sans acces a la base), utilisables par les classes metier
' et par les formulaires.
'==============================================================================

' Telephone : de 6 a 15 chiffres ; espaces, +, -, . et parentheses autorises
Public Function IsValidPhone(ByVal pText As String) As Boolean
    Dim i As Long
    Dim c As String
    Dim lngDigits As Long

    For i = 1 To Len(pText)
        c = Mid$(pText, i, 1)
        If c >= "0" And c <= "9" Then
            lngDigits = lngDigits + 1
        ElseIf InStr(" +-.()", c) = 0 Then
            Exit Function
        End If
    Next i
    IsValidPhone = (lngDigits >= 6 And lngDigits <= 15)
End Function

' Texte d'une valeur lue dans la base (Null ou vide -> chaine vide, espaces retires)
Public Function TextOfValue(ByVal pValue As Variant) As String
    If IsNull(pValue) Or IsEmpty(pValue) Then
        TextOfValue = ""
    Else
        TextOfValue = Trim$(CStr(pValue))
    End If
End Function

' Lit un montant saisi ("3 500 000", "3500000", "1 250,50") ; espaces et espaces insecables
' ignores. Retourne Faux si le texte n'est pas un nombre positif ou nul. Vide -> 0.
Public Function TryParseAmount(ByVal pText As String, ByRef pAmount As Currency) As Boolean
    Dim strText As String

    strText = Replace(Replace(Trim$(pText), " ", ""), ChrW(160), "")
    strText = Replace(strText, ChrW(8239), "")
    If Len(strText) = 0 Then
        pAmount = 0
        TryParseAmount = True
        Exit Function
    End If
    If Not IsNumeric(strText) Then Exit Function
    If CDbl(strText) < 0 Then Exit Function
    pAmount = CCur(strText)
    TryParseAmount = True
End Function
