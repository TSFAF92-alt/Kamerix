Attribute VB_Name = "modKamerixReferences"
Option Explicit
Option Private Module
'==============================================================
' Gestionnaire universel des références VBA
' Prérequis :
'   - Activer "Accès approuvé au modèle d'objet du projet VBA"
'   - Référence : Microsoft Visual Basic For Applications Extensibility 5.3
'==============================================================

Public Sub EnsureProjectReferences()

    Dim vbProj As VBIDE.VBProject
    Dim refs As Variant
    Dim i As Long

    ' Projet Kamerix lui-meme (et non le projet selectionne dans l'editeur VBA).
    ' Sans "Acces approuve au modele d'objet du projet VBA" (reglage par defaut
    ' chez un client), l'acces est refuse : on sort sans erreur, les references
    ' deja enregistrees dans le classeur restent utilisees.
    On Error Resume Next
    Set vbProj = ThisWorkbook.VBProject
    If Err.Number <> 0 Or vbProj Is Nothing Then
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0

    refs = Array( _
        Array("Visual Basic For Applications", "{000204EF-0000-0000-C000-000000000046}", 4, 2), _
        Array("OLE Automation", "{00020430-0000-0000-C000-000000000046}", 2, 0), _
        Array("Microsoft Excel Object Library", "{00020813-0000-0000-C000-000000000046}", 1, 9), _
        Array("Microsoft Office Object Library", "{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}", 2, 8), _
        Array("Microsoft Shell Controls And Automation", "{50A7E9B0-70EF-11D1-B75A-00A0C90564FE}", 1, 0), _
        Array("Microsoft Forms 2.0 Object Library", "{0D452EE1-E08F-101A-852E-02608C4D0BB4}", 2, 0), _
        Array("Microsoft Windows Common Controls 6.0 (SP6)", "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}", 2, 2), _
        Array("Microsoft VBScript Regular Expressions 5.5", "{3F4DACA7-160D-11D2-A8E9-00104B365C9F}", 5, 5), _
        Array("Microsoft ActiveX Data Objects 6.1 Library", "{B691E011-1797-432E-907A-4D8C69339129}", 6, 1), _
        Array("Microsoft Scripting Runtime", "{420B2830-E718-11CF-893D-00A0C9054228}", 1, 0), _
        Array("Windows Script Host Object Model", "{F935DC20-1CF0-11D0-ADB9-00C04FD58A0B}", 1, 0), _
        Array("Microsoft ADO Ext. 2.8 for DDL and Security", "{00000600-0000-0010-8000-00AA006D2EA4}", 2, 8), _
        Array("Microsoft Office Access Database Engine Object Library", "{4AC9E1DA-5BAD-4AC7-86E3-24F4CDCECA28}", 12, 0) _
    )

    'Debug.Print String(70, "=")
    'Debug.Print "Vérification des références VBA..."
    'Debug.Print String(70, "=")

    For i = LBound(refs) To UBound(refs)
        AddReferenceIfMissing _
            vbProj, _
            CStr(refs(i)(0)), _
            CStr(refs(i)(1)), _
            CLng(refs(i)(2)), _
            CLng(refs(i)(3))
    Next i

    'Debug.Print String(70, "=")
    'Debug.Print "Fin de la vérification."
    'Debug.Print String(70, "=")
    'MsgBox "Toutes les références ont été vérifiées.", vbInformation
End Sub

'==============================================================
' Ajoute une référence si elle n'existe pas
'==============================================================
Private Sub AddReferenceIfMissing( _
    ByVal vbProj As VBIDE.VBProject, _
    ByVal RefName As String, _
    ByVal RefGUID As String, _
    ByVal MajorVersion As Long, _
    ByVal MinorVersion As Long)

    If ReferenceExists(vbProj, RefGUID) Then
        'Debug.Print "OK     : " & RefName
        Exit Sub
    End If

    On Error GoTo ErrHandler
    vbProj.References.AddFromGuid _
        RefGUID, _
        MajorVersion, _
        MinorVersion
    'Debug.Print "AJOUTÉ : " & RefName
    Exit Sub
ErrHandler:
    MsgBox "ERREUR : " _
                & RefName _
                & " | " _
                & Err.Number _
                & " - " _
                & Err.Description, vbCritical, "Kamerix-Erreur d'Exécution"
    Err.Clear

End Sub

'==============================================================
' Vérifie si une référence existe déjà
'==============================================================
Private Function ReferenceExists(ByVal vbProj As VBIDE.VBProject, ByVal RefGUID As String) As Boolean
    Dim ref As VBIDE.Reference
    For Each ref In vbProj.References
        If StrComp(ref.GUID, RefGUID, vbTextCompare) = 0 Then
            ReferenceExists = True
            Exit Function
        End If
    Next ref
End Function

Public Sub InitializeEnvironment()
    EnsureProjectReferences
End Sub
