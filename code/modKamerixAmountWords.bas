Attribute VB_Name = "modKamerixAmountWords"
Option Explicit
Option Private Module

'==============================================================================
' KAMERIX - Montants en toutes lettres (orthographe traditionnelle)
'------------------------------------------------------------------------------
' Role        : écrire un montant en lettres sur les factures et les reçus :
'                 AmountInWords(5000000)  -> "cinq millions de francs CFA"
'                 AmountInWords(2440000)  -> "deux millions quatre cent quarante mille francs CFA"
' Orthographe : traditionnelle (décidée le 31/10/2026) :
'               - trait d'union entre dizaines et unités sous cent seulement :
'                 "vingt et un", "trente-deux", "soixante et onze", "quatre-vingt-un" ;
'               - "vingt" et "cent" prennent un s quand ils sont multipliés et terminent
'                 le nombre ("quatre-vingts", "deux cents"), mais pas devant "mille"
'                 ("quatre-vingt mille", "deux cent mille") ; devant "million(s)" et
'                 "milliard(s)", qui sont des noms, le s reste ("deux cents millions") ;
'               - "mille" est invariable et jamais précédé de "un" ;
'               - "de" devant la devise après un nombre rond de millions ou de
'                 milliards : "un million de francs CFA", mais "un million cinq cent
'                 mille francs CFA".
' Limites     : de 0 à 999 999 999 999 (moins de mille milliards) ; montant négatif
'               refusé. Partie décimale (devise à décimales) : "et NN centimes".
' Devise      : FCFA -> "franc(s) CFA" ; autre code : le code tel quel (paramètre DEVISE).
' Dependances : modKamerixSettings (CurrencyCode, CurrencyDecimals) pour AmountInWords.
'==============================================================================

Public Const ERR_AMOUNT_WORDS As Long = vbObjectError + 1910

Private Const MAX_AMOUNT As Double = 999999999999#

'==============================================================================
' API PUBLIQUE
'==============================================================================

' Montant en lettres suivi de la devise du paramètre DEVISE (minuscules)
Public Function AmountInWords(ByVal pAmount As Currency) As String
    AmountInWords = AmountInWordsFor(pAmount, CurrencyCode(), CurrencyDecimals())
End Function

' Même chose avec une devise et un nombre de décimales donnés (testable sans base)
Public Function AmountInWordsFor(ByVal pAmount As Currency, ByVal pCurrencyCode As String, _
                                 ByVal pDecimals As Long) As String
    Dim curWhole As Currency
    Dim lngCents As Long
    Dim strText As String

    If pAmount < 0 Then
        Err.Raise ERR_AMOUNT_WORDS, "modKamerixAmountWords.AmountInWordsFor", _
            "Montant négatif : il ne peut pas être écrit en lettres."
    End If

    curWhole = Int(pAmount)
    If pDecimals > 0 Then lngCents = CLng((pAmount - curWhole) * 100)

    strText = NumberInWords(curWhole)
    If EndsWithMillionOrBillion(curWhole) Then strText = strText & " de"
    strText = strText & " " & CurrencyWords(pCurrencyCode, curWhole >= 2)
    If lngCents > 0 Then strText = strText & " et " & NumberInWords(lngCents) & " centime" & IIf(lngCents >= 2, "s", "")

    AmountInWordsFor = strText
End Function

' Nombre entier en lettres, sans devise : 80 -> "quatre-vingts", 21000 -> "vingt et un mille"
Public Function NumberInWords(ByVal pNumber As Currency) As String
    Dim curValue As Currency
    Dim lngBillions As Long
    Dim lngMillions As Long
    Dim lngThousands As Long
    Dim lngRest As Long
    Dim strText As String

    If pNumber < 0 Or pNumber > MAX_AMOUNT Or pNumber <> Int(pNumber) Then
        Err.Raise ERR_AMOUNT_WORDS, "modKamerixAmountWords.NumberInWords", _
            "Nombre impossible à écrire en lettres : " & CStr(pNumber) & _
            " (entier de 0 à 999 999 999 999)."
    End If
    If pNumber = 0 Then
        NumberInWords = "zéro"
        Exit Function
    End If

    curValue = pNumber
    lngBillions = CLng(Int(curValue / 1000000000))
    curValue = curValue - CCur(lngBillions) * 1000000000
    lngMillions = CLng(Int(curValue / 1000000))
    curValue = curValue - CCur(lngMillions) * 1000000
    lngThousands = CLng(Int(curValue / 1000))
    lngRest = CLng(curValue - CCur(lngThousands) * 1000)

    If lngBillions > 0 Then
        strText = JoinWords(strText, Below1000(lngBillions, True) & " milliard" & IIf(lngBillions >= 2, "s", ""))
    End If
    If lngMillions > 0 Then
        strText = JoinWords(strText, Below1000(lngMillions, True) & " million" & IIf(lngMillions >= 2, "s", ""))
    End If
    If lngThousands = 1 Then
        strText = JoinWords(strText, "mille")
    ElseIf lngThousands > 1 Then
        strText = JoinWords(strText, Below1000(lngThousands, False) & " mille")
    End If
    If lngRest > 0 Then strText = JoinWords(strText, Below1000(lngRest, True))

    NumberInWords = strText
End Function

' Première lettre en majuscule ("Cinq millions de francs CFA")
Public Function CapitalizeFirst(ByVal pText As String) As String
    If Len(pText) = 0 Then Exit Function
    CapitalizeFirst = UCase$(Left$(pText, 1)) & Mid$(pText, 2)
End Function

'==============================================================================
' PROCEDURES INTERNES
'==============================================================================

' 1 à 999. pEndsNumber : Faux si le groupe est suivi de "mille" (cent et vingt restent
' alors sans s : "deux cent mille", "quatre-vingt mille").
Private Function Below1000(ByVal pNumber As Long, ByVal pEndsNumber As Boolean) As String
    Dim lngHundreds As Long
    Dim lngRest As Long
    Dim strText As String

    lngHundreds = pNumber \ 100
    lngRest = pNumber Mod 100

    Select Case lngHundreds
        Case 0
            strText = Below100(lngRest)
        Case 1
            strText = "cent"
            If lngRest > 0 Then strText = strText & " " & Below100(lngRest)
        Case Else
            strText = UnitWord(lngHundreds) & " cent"
            If lngRest > 0 Then
                strText = strText & " " & Below100(lngRest)
            ElseIf pEndsNumber Then
                strText = strText & "s"
            End If
    End Select

    ' "quatre-vingts" perd son s devant "mille"
    If Not pEndsNumber And Right$(strText, 13) = "quatre-vingts" Then strText = Left$(strText, Len(strText) - 1)

    Below1000 = strText
End Function

' 1 à 99
Private Function Below100(ByVal pNumber As Long) As String
    Dim lngTens As Long
    Dim lngUnits As Long

    If pNumber < 17 Then
        Below100 = UnitWord(pNumber)
        Exit Function
    End If
    If pNumber < 20 Then
        Below100 = "dix-" & UnitWord(pNumber - 10)
        Exit Function
    End If

    lngTens = pNumber \ 10
    lngUnits = pNumber Mod 10

    Select Case lngTens
        Case 7                                    ' soixante-dix à soixante-dix-neuf
            If lngUnits = 1 Then
                Below100 = "soixante et onze"
            Else
                Below100 = "soixante-" & Below100(10 + lngUnits)
            End If
        Case 8                                    ' quatre-vingts à quatre-vingt-neuf
            If lngUnits = 0 Then
                Below100 = "quatre-vingts"
            Else
                Below100 = "quatre-vingt-" & UnitWord(lngUnits)
            End If
        Case 9                                    ' quatre-vingt-dix à quatre-vingt-dix-neuf
            Below100 = "quatre-vingt-" & Below100(10 + lngUnits)
        Case Else                                 ' vingt à soixante-neuf
            If lngUnits = 0 Then
                Below100 = TensWord(lngTens)
            ElseIf lngUnits = 1 Then
                Below100 = TensWord(lngTens) & " et un"
            Else
                Below100 = TensWord(lngTens) & "-" & UnitWord(lngUnits)
            End If
    End Select
End Function

' 0 à 16
Private Function UnitWord(ByVal pNumber As Long) As String
    UnitWord = Choose(pNumber + 1, "zéro", "un", "deux", "trois", "quatre", "cinq", "six", "sept", "huit", "neuf", _
                      "dix", "onze", "douze", "treize", "quatorze", "quinze", "seize")
End Function

' 2 à 6 (vingt à soixante)
Private Function TensWord(ByVal pTens As Long) As String
    TensWord = Choose(pTens - 1, "vingt", "trente", "quarante", "cinquante", "soixante")
End Function

Private Function JoinWords(ByVal pLeft As String, ByVal pRight As String) As String
    If Len(pLeft) = 0 Then JoinWords = pRight Else JoinWords = pLeft & " " & pRight
End Function

' Vrai pour un nombre rond de millions ou de milliards (1 000 000, 3 000 000 000...)
Private Function EndsWithMillionOrBillion(ByVal pNumber As Currency) As Boolean
    If pNumber < 1000000 Then Exit Function
    EndsWithMillionOrBillion = (pNumber - Int(pNumber / 1000000) * 1000000 = 0)
End Function

' Nom de la devise : FCFA -> franc(s) CFA ; autre code : tel quel
Private Function CurrencyWords(ByVal pCode As String, ByVal pPlural As Boolean) As String
    Select Case UCase$(Trim$(pCode))
        Case "FCFA", "XAF", "XOF", "CFA"
            CurrencyWords = IIf(pPlural, "francs CFA", "franc CFA")
        Case Else
            CurrencyWords = Trim$(pCode)
    End Select
End Function
