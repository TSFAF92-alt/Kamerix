Attribute VB_Name = "ZmodulPictureGDIPlus"
Option Explicit
Option Private Module

'==============================================================================
' Chargement d'images PNG, JPG, BMP, GIF, TIFF via GDI+ (module de l'utilisateur,
' revu le 01/11/2026 pour Kamerix)
'------------------------------------------------------------------------------
' LoadPictureGDIPlus(chemin [, couleur de fond]) retourne une IPicture utilisable
' par le ruban (getImage) et par les contrôles Image des userforms.
' Retourne Nothing si l'image ne peut pas être lue.
' Couleur de fond : remplace la transparence (ARGB). 0 par défaut (noir transparent,
' comportement d'origine) ; GDIP_WHITE_BACKGROUND pour un fond blanc (logo d'un formulaire).
' Corrections du 01/11/2026 :
'   - MsgBox du gestionnaire d'erreur : le titre était passé à la place des boutons
'     (erreur « Type incompatible » au lieu du message) ;
'   - GDI+ est toujours arrêté et le bitmap toujours libéré, y compris en cas d'erreur ;
'   - paramètre facultatif de couleur de fond (PNG transparents affichés sur fond noir).
' Compatible VBA 32 et 64 bits (VBA7).
'==============================================================================

#If VBA7 Then
    Private Declare PtrSafe Function GdiplusStartup Lib "gdiplus" (ByRef token As LongPtr, ByRef inputbuf As GDIPLUSSTARTUPINPUT, ByVal outputbuf As LongPtr) As Long
    Private Declare PtrSafe Sub GdiplusShutdown Lib "gdiplus" (ByVal token As LongPtr)
    Private Declare PtrSafe Function GdipCreateBitmapFromFile Lib "gdiplus" (ByVal fileName As LongPtr, ByRef bitmap As LongPtr) As Long
    Private Declare PtrSafe Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" (ByVal nbitmap As LongPtr, ByRef hbmReturn As LongPtr, ByVal background As Long) As Long
    Private Declare PtrSafe Function GdipDisposeImage Lib "gdiplus" (ByVal nimage As LongPtr) As Long
    Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32" (ByRef PicDesc As PICTDESC, ByRef RefIID As GUID_TYPE, ByVal fOwn As Long, ByRef Ipic As IPicture) As Long
    Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
#Else
    Private Declare Function GdiplusStartup Lib "gdiplus" (ByRef token As Long, ByRef inputbuf As GDIPLUSSTARTUPINPUT, ByVal outputbuf As Long) As Long
    Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal token As Long)
    Private Declare Function GdipCreateBitmapFromFile Lib "gdiplus" (ByVal fileName As Long, ByRef bitmap As Long) As Long
    Private Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" (ByVal nbitmap As Long, ByRef hbmReturn As Long, ByVal background As Long) As Long
    Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal nimage As Long) As Long
    Private Declare Function OleCreatePictureIndirect Lib "oleaut32" (ByRef PicDesc As PICTDESC, ByRef RefIID As GUID_TYPE, ByVal fOwn As Long, ByRef Ipic As IPicture) As Long
    Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
#End If

Private Type GDIPLUSSTARTUPINPUT
    GdiplusVersion As Long
    #If VBA7 Then
        DebugEventCallback As LongPtr
    #Else
        DebugEventCallback As Long
    #End If
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Type PICTDESC
    cbSizeOfStruct As Long
    picType As Long
    #If VBA7 Then
        hImage As LongPtr
        hPal As LongPtr
    #Else
        hImage As Long
        hPal As Long
    #End If
End Type

Private Type GUID_TYPE
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Private Const PICTYPE_BITMAP As Long = 1

' Fond blanc opaque (ARGB FFFFFFFF) pour remplacer la transparence
Public Const GDIP_WHITE_BACKGROUND As Long = &HFFFFFFFF

' Charge un fichier image via GDI+ et le retourne comme IPicture ; Nothing en cas d'échec.
' pBackColor : couleur ARGB qui remplace la transparence (0 = comportement d'origine).
Public Function LoadPictureGDIPlus(ByVal sFile As String, Optional ByVal pBackColor As Long = 0) As IPicture
    #If VBA7 Then
        Dim gdipToken As LongPtr
        Dim gdipBitmap As LongPtr
        Dim hBitmap As LongPtr
    #Else
        Dim gdipToken As Long
        Dim gdipBitmap As Long
        Dim hBitmap As Long
    #End If
    Dim gdipInput As GDIPLUSSTARTUPINPUT
    Dim Ipic As PICTDESC
    Dim IidPicture As GUID_TYPE
    Dim resultPic As IPicture
    Dim strError As String

    On Error GoTo EH

    gdipInput.GdiplusVersion = 1
    If GdiplusStartup(gdipToken, gdipInput, 0) <> 0 Then Exit Function

    If GdipCreateBitmapFromFile(StrPtr(sFile), gdipBitmap) = 0 Then
        If GdipCreateHBITMAPFromBitmap(gdipBitmap, hBitmap, pBackColor) = 0 Then
            ' IID_IPicture = {7BF80980-BF32-101A-8BBB-00AA00300CAB}
            With IidPicture
                .Data1 = &H7BF80980
                .Data2 = &HBF32
                .Data3 = &H101A
                .Data4(0) = &H8B: .Data4(1) = &HBB: .Data4(2) = &H0: .Data4(3) = &HAA
                .Data4(4) = &H0: .Data4(5) = &H30: .Data4(6) = &HC: .Data4(7) = &HAB
            End With
            With Ipic
                .cbSizeOfStruct = LenB(Ipic)
                .picType = PICTYPE_BITMAP
                .hImage = hBitmap
                .hPal = 0
            End With
            ' fOwn = 1 : l'image devient propriétaire du bitmap (libéré avec elle)
            If OleCreatePictureIndirect(Ipic, IidPicture, 1, resultPic) = 0 Then
                Set LoadPictureGDIPlus = resultPic
                hBitmap = 0                       ' ne plus le libérer ici : il appartient à l'image
            End If
        End If
    End If

Cleanup:
    On Error Resume Next
    If hBitmap <> 0 Then DeleteObject hBitmap
    If gdipBitmap <> 0 Then GdipDisposeImage gdipBitmap
    If gdipToken <> 0 Then GdiplusShutdown gdipToken
    On Error GoTo 0
    If Len(strError) > 0 Then MsgBox strError, vbExclamation, "Kamerix"
    Exit Function

EH:
    strError = "Erreur GDI+ sur " & sFile & " : " & Err.Description
    Resume Cleanup
End Function
