Attribute VB_Name = "modKamerixCrypto"
Option Explicit
Option Private Module

'==============================================================================
' KAMERIX - Hash des mots de passe : PBKDF2-HMAC-SHA256 via Windows CNG (bcrypt.dll)
'------------------------------------------------------------------------------
' Principe    : le mot de passe n'est jamais stocke. On stocke un sel aleatoire
'               (Utilisateur.Sel) et le resultat de PBKDF2-HMAC-SHA256 avec un
'               grand nombre d'iterations (Utilisateur.MotDePasseHash), sous la
'               forme  "iterations$hash_hexadecimal"  (ex. 600000$ab12...).
'               Le nombre d'iterations est dans la valeur stockee : on pourra
'               l'augmenter plus tard sans casser les anciens comptes.
' Pourquoi    : un hash lent (600 000 iterations) rend tres couteux le calcul
'               d'un grand nombre d'essais si quelqu'un obtient le fichier .accdb.
' Aleatoire   : BCryptGenRandom (generateur cryptographique du systeme).
' Prerequis   : Windows 10 ou plus recent, Office 2010 ou plus (32 ou 64 bits).
' Erreurs     : remontees avec Err.Raise (ERR_CRYPTO). Aucune MsgBox, sauf dans
'               la procedure de test TestCrypto.
' API         : GenerateSalt, HashPassword, VerifyPassword, NeedsRehash,
'               DeriveKey, RandomBytes, BytesToHex, HexToBytes.
'==============================================================================

Public Const ERR_CRYPTO As Long = ERR_BASE + 10

Public Const DEFAULT_ITERATIONS As Long = 310000 '600000
Private Const MIN_ITERATIONS_VERIFY As Long = 1000
Private Const MAX_ITERATIONS_VERIFY As Long = 10000000
Private Const SALT_BYTES As Long = 32          ' 32 octets = 64 caracteres hexa (Utilisateur.Sel)
Private Const HASH_BYTES As Long = 32          ' 32 octets = 64 caracteres hexa

Private Const BCRYPT_ALG_HANDLE_HMAC_FLAG As Long = &H8
Private Const BCRYPT_USE_SYSTEM_PREFERRED_RNG As Long = &H2
Private Const CP_UTF8 As Long = 65001

#If VBA7 Then
    Private Declare PtrSafe Function BCryptOpenAlgorithmProvider Lib "bcrypt.dll" ( _
        ByRef phAlgorithm As LongPtr, ByVal pszAlgId As LongPtr, _
        ByVal pszImplementation As LongPtr, ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function BCryptCloseAlgorithmProvider Lib "bcrypt.dll" ( _
        ByVal hAlgorithm As LongPtr, ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function BCryptGenRandom Lib "bcrypt.dll" ( _
        ByVal hAlgorithm As LongPtr, ByVal pbBuffer As LongPtr, _
        ByVal cbBuffer As Long, ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" ( _
        ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, _
        ByVal cchWideChar As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, _
        ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As LongPtr) As Long
    ' cIterations est un entier 64 bits : un seul argument en 64 bits, deux Long en 32 bits
    #If Win64 Then
        Private Declare PtrSafe Function BCryptDeriveKeyPBKDF2 Lib "bcrypt.dll" ( _
            ByVal hPrf As LongPtr, ByVal pbPassword As LongPtr, ByVal cbPassword As Long, _
            ByVal pbSalt As LongPtr, ByVal cbSalt As Long, ByVal cIterations As LongLong, _
            ByVal pbDerivedKey As LongPtr, ByVal cbDerivedKey As Long, ByVal dwFlags As Long) As Long
    #Else
        Private Declare PtrSafe Function BCryptDeriveKeyPBKDF2 Lib "bcrypt.dll" ( _
            ByVal hPrf As LongPtr, ByVal pbPassword As LongPtr, ByVal cbPassword As Long, _
            ByVal pbSalt As LongPtr, ByVal cbSalt As Long, ByVal cIterLow As Long, ByVal cIterHigh As Long, _
            ByVal pbDerivedKey As LongPtr, ByVal cbDerivedKey As Long, ByVal dwFlags As Long) As Long
    #End If
#End If

'==============================================================================
' API PUBLIQUE
'==============================================================================

' Sel aleatoire de 64 caracteres hexadecimaux, a stocker dans Utilisateur.Sel
Public Function GenerateSalt() As String
    Dim bytSalt() As Byte

    bytSalt = RandomBytes(SALT_BYTES)
    GenerateSalt = BytesToHex(bytSalt)
End Function

' Calcule la valeur a stocker dans Utilisateur.MotDePasseHash : "iterations$hash"
Public Function HashPassword(ByVal strPassword As String, ByVal strSaltHex As String, _
                             Optional ByVal lngIterations As Long = DEFAULT_ITERATIONS) As String
    Dim bytPwd() As Byte
    Dim bytSalt() As Byte
    Dim bytKey() As Byte

    If Len(strPassword) = 0 Then
        Err.Raise ERR_CRYPTO, "modKamerixCrypto.HashPassword", "Le mot de passe est vide."
    End If

    bytPwd = ToUtf8(strPassword)
    bytSalt = HexToBytes(strSaltHex)
    bytKey = DeriveKey(bytPwd, bytSalt, lngIterations, HASH_BYTES)

    HashPassword = CStr(lngIterations) & "$" & BytesToHex(bytKey)

    Wipe bytPwd
    Wipe bytKey
End Function

' Vrai si le mot de passe correspond a la valeur stockee. Ne leve pas d'erreur
' pour une valeur stockee mal formee ou un mot de passe vide : retourne False.
Public Function VerifyPassword(ByVal strPassword As String, ByVal strSaltHex As String, _
                               ByVal strStored As String) As Boolean
    Dim lngPos As Long
    Dim lngIter As Long
    Dim strExpected As String
    Dim strActual As String
    Dim bytPwd() As Byte
    Dim bytSalt() As Byte
    Dim bytKey() As Byte

    If Len(strPassword) = 0 Then Exit Function

    lngPos = InStr(strStored, "$")
    If lngPos < 2 Then Exit Function
    If Not IsNumeric(Left$(strStored, lngPos - 1)) Then Exit Function
    lngIter = CLng(Left$(strStored, lngPos - 1))
    If lngIter < MIN_ITERATIONS_VERIFY Or lngIter > MAX_ITERATIONS_VERIFY Then Exit Function
    strExpected = LCase$(Mid$(strStored, lngPos + 1))

    bytPwd = ToUtf8(strPassword)
    bytSalt = HexToBytes(strSaltHex)
    bytKey = DeriveKey(bytPwd, bytSalt, lngIter, HASH_BYTES)
    strActual = BytesToHex(bytKey)

    VerifyPassword = ConstantTimeEquals(strActual, strExpected)

    Wipe bytPwd
    Wipe bytKey
End Function

' Vrai si la valeur stockee utilise moins d'iterations que le standard actuel
' (a utiliser apres une connexion reussie pour remettre le hash a niveau)
Public Function NeedsRehash(ByVal strStored As String) As Boolean
    Dim lngPos As Long

    lngPos = InStr(strStored, "$")
    If lngPos < 2 Then
        NeedsRehash = True
    ElseIf Not IsNumeric(Left$(strStored, lngPos - 1)) Then
        NeedsRehash = True
    Else
        NeedsRehash = (CLng(Left$(strStored, lngPos - 1)) < DEFAULT_ITERATIONS)
    End If
End Function

' PBKDF2-HMAC-SHA256. Retourne lngLen octets derives du mot de passe et du sel.
' Les tableaux doivent etre non vides (indices a partir de 0 ou de tout autre LBound).
Public Function DeriveKey(ByRef bytPwd() As Byte, ByRef bytSalt() As Byte, _
                          ByVal lngIter As Long, ByVal lngLen As Long) As Byte()
    Dim hAlg As LongPtr
    Dim lngStatus As Long
    Dim bytOut() As Byte
    Dim strAlg As String
    Dim lngPwdLen As Long
    Dim lngSaltLen As Long
    Dim lngErr As Long
    Dim strErr As String
    Dim strSrc As String

    If lngIter < 1 Or lngLen < 1 Then
        Err.Raise ERR_CRYPTO, "modKamerixCrypto.DeriveKey", "Parametres invalides (iterations ou longueur)."
    End If

    lngPwdLen = UBound(bytPwd) - LBound(bytPwd) + 1
    lngSaltLen = UBound(bytSalt) - LBound(bytSalt) + 1
    If lngPwdLen < 1 Or lngSaltLen < 1 Then
        Err.Raise ERR_CRYPTO, "modKamerixCrypto.DeriveKey", "Mot de passe ou sel vide."
    End If

    ReDim bytOut(0 To lngLen - 1)
    strAlg = "SHA256"

    lngStatus = BCryptOpenAlgorithmProvider(hAlg, StrPtr(strAlg), 0, BCRYPT_ALG_HANDLE_HMAC_FLAG)
    If lngStatus <> 0 Then
        Err.Raise ERR_CRYPTO, "modKamerixCrypto.DeriveKey", _
            "Ouverture du fournisseur CNG impossible (code 0x" & Hex$(lngStatus) & "). " & _
            "Windows 10 ou plus recent est requis."
    End If

    On Error GoTo EH

    #If Win64 Then
        lngStatus = BCryptDeriveKeyPBKDF2(hAlg, VarPtr(bytPwd(LBound(bytPwd))), lngPwdLen, _
            VarPtr(bytSalt(LBound(bytSalt))), lngSaltLen, CLngLng(lngIter), _
            VarPtr(bytOut(0)), lngLen, 0)
    #Else
        lngStatus = BCryptDeriveKeyPBKDF2(hAlg, VarPtr(bytPwd(LBound(bytPwd))), lngPwdLen, _
            VarPtr(bytSalt(LBound(bytSalt))), lngSaltLen, lngIter, 0, _
            VarPtr(bytOut(0)), lngLen, 0)
    #End If

    If lngStatus <> 0 Then
        Err.Raise ERR_CRYPTO, "modKamerixCrypto.DeriveKey", _
            "Derivation PBKDF2 impossible (code 0x" & Hex$(lngStatus) & ")."
    End If

    BCryptCloseAlgorithmProvider hAlg, 0
    DeriveKey = bytOut
    Exit Function

EH:
    lngErr = Err.Number
    strErr = Err.Description
    strSrc = Err.Source

    On Error Resume Next
    BCryptCloseAlgorithmProvider hAlg, 0
    On Error GoTo 0

    Err.Raise lngErr, strSrc, strErr
End Function

' Octets aleatoires cryptographiques (generateur du systeme)
Public Function RandomBytes(ByVal lngCount As Long) As Byte()
    Dim bytOut() As Byte
    Dim lngStatus As Long

    If lngCount < 1 Then
        Err.Raise ERR_CRYPTO, "modKamerixCrypto.RandomBytes", "Nombre d'octets invalide."
    End If

    ReDim bytOut(0 To lngCount - 1)
    lngStatus = BCryptGenRandom(0, VarPtr(bytOut(0)), lngCount, BCRYPT_USE_SYSTEM_PREFERRED_RNG)
    If lngStatus <> 0 Then
        Err.Raise ERR_CRYPTO, "modKamerixCrypto.RandomBytes", _
            "Generation aleatoire impossible (code 0x" & Hex$(lngStatus) & ")."
    End If

    RandomBytes = bytOut
End Function

' Octets -> hexadecimal minuscule
Public Function BytesToHex(ByRef bytData() As Byte) As String
    Dim i As Long
    Dim strOut As String

    For i = LBound(bytData) To UBound(bytData)
        strOut = strOut & Right$("0" & LCase$(Hex$(bytData(i))), 2)
    Next i
    BytesToHex = strOut
End Function

' Hexadecimal -> octets (erreur si longueur impaire ou caractere non hexadecimal)
Public Function HexToBytes(ByVal strHex As String) As Byte()
    Dim bytOut() As Byte
    Dim i As Long
    Dim strPair As String

    If Len(strHex) = 0 Or (Len(strHex) Mod 2) <> 0 Then
        Err.Raise ERR_CRYPTO, "modKamerixCrypto.HexToBytes", "Chaine hexadecimale invalide (longueur)."
    End If

    ReDim bytOut(0 To Len(strHex) \ 2 - 1)
    For i = 0 To UBound(bytOut)
        strPair = Mid$(strHex, i * 2 + 1, 2)
        If Not (strPair Like "[0-9A-Fa-f][0-9A-Fa-f]") Then
            Err.Raise ERR_CRYPTO, "modKamerixCrypto.HexToBytes", "Chaine hexadecimale invalide (caractere)."
        End If
        bytOut(i) = CByte(Val("&H" & strPair))
    Next i

    HexToBytes = bytOut
End Function

'==============================================================================
' PROCEDURES INTERNES
'==============================================================================

' Texte -> octets UTF-8 (les accents du mot de passe sont geres)
Private Function ToUtf8(ByVal strText As String) As Byte()
    Dim lngLen As Long
    Dim bytOut() As Byte

    lngLen = WideCharToMultiByte(CP_UTF8, 0, StrPtr(strText), Len(strText), 0, 0, 0, 0)
    If lngLen <= 0 Then
        Err.Raise ERR_CRYPTO, "modKamerixCrypto.ToUtf8", "Conversion UTF-8 impossible."
    End If

    ReDim bytOut(0 To lngLen - 1)
    lngLen = WideCharToMultiByte(CP_UTF8, 0, StrPtr(strText), Len(strText), _
                                 VarPtr(bytOut(0)), lngLen, 0, 0)
    If lngLen <= 0 Then
        Err.Raise ERR_CRYPTO, "modKamerixCrypto.ToUtf8", "Conversion UTF-8 impossible."
    End If

    ToUtf8 = bytOut
End Function

' Comparaison en temps constant (ne s'arrete pas au premier caractere different)
Private Function ConstantTimeEquals(ByVal strA As String, ByVal strB As String) As Boolean
    Dim i As Long
    Dim lngDiff As Long

    If Len(strA) <> Len(strB) Then Exit Function

    For i = 1 To Len(strA)
        lngDiff = lngDiff Or (AscW(Mid$(strA, i, 1)) Xor AscW(Mid$(strB, i, 1)))
    Next i

    ConstantTimeEquals = (lngDiff = 0)
End Function

' Remet a zero un tableau d'octets sensible
Private Sub Wipe(ByRef bytData() As Byte)
    Dim i As Long

    On Error Resume Next
    For i = LBound(bytData) To UBound(bytData)
        bytData(i) = 0
    Next i
    On Error GoTo 0
End Sub

'==============================================================================
' PROCEDURE DE TEST (a lancer avec F5 depuis l'editeur VBA)
' Les 3 premiers tests utilisent les vecteurs de test standard de
' PBKDF2-HMAC-SHA256 (mot de passe "password", sel "salt").
' En cas d'echec, le rapport affiche la valeur obtenue et la valeur attendue.
'==============================================================================
Public Sub TestCrypto()
    Dim lngOk As Long
    Dim lngTotal As Long
    Dim strReport As String
    Dim bytPwd() As Byte
    Dim bytSalt() As Byte
    Dim bytKey() As Byte
    Dim strSalt1 As String
    Dim strSalt2 As String
    Dim strH1 As String
    Dim strH2 As String
    Dim strPwd As String
    Dim sngStart As Single
    Dim sngElapsed As Single

    On Error GoTo EH

    bytPwd = ToUtf8("password")
    bytSalt = ToUtf8("salt")

    strReport = "Octets mot de passe : " & BytesToHex(bytPwd) & " (attendu 70617373776f7264)" & vbCrLf & _
                 "Octets sel : " & BytesToHex(bytSalt) & " (attendu 73616c74)" & vbCrLf & vbCrLf

    bytKey = DeriveKey(bytPwd, bytSalt, 1, 32)
    CheckEq "Vecteur standard, 1 iteration", BytesToHex(bytKey), _
        "120fb6cffcf8b32c43e7225256c4f837a86548c92ccc35480805987cb70be17b", _
        lngOk, lngTotal, strReport

    bytKey = DeriveKey(bytPwd, bytSalt, 2, 32)
    CheckEq "Vecteur standard, 2 iterations", BytesToHex(bytKey), _
        "ae4d0c95af6b46d32d0adff928f06dd02a303f8ef3c251dfd6e2d85a95474c43", _
        lngOk, lngTotal, strReport

    bytKey = DeriveKey(bytPwd, bytSalt, 4096, 32)
    CheckEq "Vecteur standard, 4096 iterations", BytesToHex(bytKey), _
        "c5e478d59288c841aa530db6845c4c8d962893a001ce4e11a4963873aa98134a", _
        lngOk, lngTotal, strReport

    strSalt1 = GenerateSalt()
    strSalt2 = GenerateSalt()
    CheckTrue "Sel de 64 caracteres", (Len(strSalt1) = 64), lngOk, lngTotal, strReport
    CheckTrue "Deux sels differents", (strSalt1 <> strSalt2), lngOk, lngTotal, strReport

    ' Mot de passe avec accents et symbole euro
    strPwd = "Mot2Passe-d" & ChrW(233) & "t" & ChrW(233) & ChrW(8364)

    sngStart = Timer
    strH1 = HashPassword(strPwd, strSalt1)
    sngElapsed = Timer - sngStart

    CheckTrue "Format iterations$hash", _
        (Left$(strH1, Len(CStr(DEFAULT_ITERATIONS)) + 1) = CStr(DEFAULT_ITERATIONS) & "$"), _
        lngOk, lngTotal, strReport
    CheckTrue "Valeur stockee <= 128 caracteres", (Len(strH1) <= 128), lngOk, lngTotal, strReport
    CheckTrue "Mot de passe correct accepte", VerifyPassword(strPwd, strSalt1, strH1), _
        lngOk, lngTotal, strReport
    CheckTrue "Mauvais mot de passe refuse", Not VerifyPassword(strPwd & "x", strSalt1, strH1), _
        lngOk, lngTotal, strReport
    CheckTrue "Mauvais sel refuse", Not VerifyPassword(strPwd, strSalt2, strH1), _
        lngOk, lngTotal, strReport

    strH2 = HashPassword(strPwd, strSalt2)
    CheckTrue "Meme mot de passe, sels differents : hash differents", (strH1 <> strH2), _
        lngOk, lngTotal, strReport
    CheckTrue "Valeur stockee mal formee refusee", Not VerifyPassword(strPwd, strSalt1, "n'importe quoi"), _
        lngOk, lngTotal, strReport
    CheckTrue "Hash a niveau : pas de remise a niveau", Not NeedsRehash(strH1), _
        lngOk, lngTotal, strReport

    Debug.Print lngOk & " tests reussis sur " & lngTotal & "." & vbCrLf & strReport   ' rapport complet (Ctrl+G)
    MsgBox lngOk & " tests reussis sur " & lngTotal & "." & vbCrLf & vbCrLf & strReport & vbCrLf & _
        "Duree d'un hash a " & DEFAULT_ITERATIONS & " iterations : " & Format$(sngElapsed, "0.00") & " s", _
        IIf(lngOk = lngTotal, vbInformation, vbExclamation), "Kamerix - hash des mots de passe"
    Exit Sub

EH:
    MsgBox "Echec (erreur " & Err.Number & ") :" & vbCrLf & vbCrLf & Err.Description, vbCritical, "Kamerix"
End Sub

' Comparaison de deux textes ; en cas d'echec, affiche obtenu et attendu
Private Sub CheckEq(ByVal strLabel As String, ByVal strActual As String, ByVal strExpected As String, _
                    ByRef lngOk As Long, ByRef lngTotal As Long, ByRef strReport As String)
    lngTotal = lngTotal + 1
    If strActual = strExpected Then
        lngOk = lngOk + 1
        strReport = strReport & "OK     " & strLabel & vbCrLf
    Else
        strReport = strReport & "ECHEC  " & strLabel & vbCrLf & _
            "          obtenu  : " & Left$(strActual, 40) & vbCrLf & _
            "          attendu : " & Left$(strExpected, 40) & vbCrLf
    End If
End Sub

' Verification d'une condition vraie (independant de la langue de Windows/Office)
Private Sub CheckTrue(ByVal strLabel As String, ByVal blnCondition As Boolean, _
                      ByRef lngOk As Long, ByRef lngTotal As Long, ByRef strReport As String)
    lngTotal = lngTotal + 1
    If blnCondition Then
        lngOk = lngOk + 1
        strReport = strReport & "OK     " & strLabel & vbCrLf
    Else
        strReport = strReport & "ECHEC  " & strLabel & vbCrLf
    End If
End Sub
