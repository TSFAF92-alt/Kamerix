Attribute VB_Name = "ZmodImportTool"
Option Explicit
Option Private Module

'==============================================================================
' KAMERIX - Outil : remplacement de modules VBA a partir de fichiers exportes
'------------------------------------------------------------------------------
' ReplaceKamerixModules : choisir un ou plusieurs fichiers (.bas, .cls, .frm),
' en nombre quelconque. Pour chaque fichier, le nom du module est lu dans la
' ligne "Attribute VB_Name" du fichier (et non dans le nom du fichier) :
'   - un module de ce nom existe : il est sauvegarde (export), supprime, puis
'     le fichier est importe a sa place ;
'   - aucun module de ce nom : le fichier est importe comme nouveau module.
' Jamais remplaces : cet outil lui-meme, ThisWorkbook et les feuilles (modules
' de document : leur code se colle a la main).
' Sauvegarde  : dossier "Backup_VBA_aaaammjj_hhnnss" a cote du classeur, cree
'               seulement si au moins un module est remplace.
' Apres usage : Debogage > Compiler VBAProject, puis relancer les tests.
' Prerequis   : "Acces approuve au modele d'objet du projet VBA" coche.
' Liaison     : tardive (aucune reference VBIDE necessaire).
'==============================================================================

Private Const VBEXT_STD_MODULE As Long = 1
Private Const VBEXT_CLASS_MODULE As Long = 2
Private Const VBEXT_MS_FORM As Long = 3
Private Const VBEXT_DOCUMENT As Long = 100
Private Const VBEXT_PP_LOCKED As Long = 1
Private Const MSO_FILE_DIALOG_FILE_PICKER As Long = 3
Private Const MAX_MODULE_NAME As Long = 31

Private Const THIS_MODULE As String = "ZmodImportTool"
Private Const TOOL_TITLE As String = "Kamerix - remplacement de modules"

' Resultat du traitement d'un fichier
Private Enum KxImportResult
    kxImportReplaced = 1
    kxImportAdded = 2
    kxImportSkipped = 3
    kxImportFailed = 4
End Enum

'==============================================================================
' API PUBLIQUE
'==============================================================================

' Point d'entree (F5) : choix des fichiers, puis remplacement
Public Sub ReplaceKamerixModules()
    Dim vFiles As Variant

    vFiles = PickModuleFiles()
    If IsEmpty(vFiles) Then Exit Sub

    ReplaceModulesFromFiles vFiles
End Sub

' Remplace ou ajoute les modules decrits par une liste de chemins de fichiers.
' Utilisable aussi sans boite de dialogue :
'     ReplaceModulesFromFiles Array("C:\Kamerix\code\modKamerixSession.bas", "C:\Kamerix\code\Car.cls")
Public Sub ReplaceModulesFromFiles(ByVal vFiles As Variant)
    Dim objProject As Object
    Dim strBackupFolder As String
    Dim strLine As String
    Dim strReport As String
    Dim lngReplaced As Long
    Dim lngAdded As Long
    Dim lngSkipped As Long
    Dim lngFailed As Long
    Dim i As Long

    Set objProject = GetProject()
    If objProject Is Nothing Then Exit Sub

    For i = LBound(vFiles) To UBound(vFiles)
        Select Case ProcessFile(objProject, CStr(vFiles(i)), strBackupFolder, strLine)
            Case kxImportReplaced: lngReplaced = lngReplaced + 1
            Case kxImportAdded:    lngAdded = lngAdded + 1
            Case kxImportSkipped:  lngSkipped = lngSkipped + 1
            Case Else:             lngFailed = lngFailed + 1
        End Select
        strReport = strReport & strLine & vbCrLf
    Next i

    strReport = lngReplaced & " remplace(s), " & lngAdded & " ajoute(s), " & _
                lngSkipped & " ignore(s), " & lngFailed & " en echec." & vbCrLf & vbCrLf & strReport
    If Len(strBackupFolder) > 0 Then
        strReport = strReport & vbCrLf & "Sauvegarde des anciens modules :" & vbCrLf & strBackupFolder & vbCrLf
    End If
    strReport = strReport & vbCrLf & "Etape suivante : Debogage > Compiler VBAProject, puis relancer les tests."

    Debug.Print strReport
    MsgBox strReport, IIf(lngFailed = 0, vbInformation, vbExclamation), TOOL_TITLE
End Sub

'==============================================================================
' PROCEDURES INTERNES
'==============================================================================

' Traite un fichier. strLine recoit la ligne du rapport.
' strBackupFolder est cree au premier remplacement puis reutilise.
Private Function ProcessFile(ByVal objProject As Object, ByVal strFile As String, _
                             ByRef strBackupFolder As String, ByRef strLine As String) As KxImportResult
    Dim objOld As Object
    Dim objNew As Object
    Dim strName As String
    Dim strBackupFile As String
    Dim blnRemoved As Boolean

    On Error GoTo EH

    strLine = "?      " & FileNameOf(strFile)

    If Len(Dir$(strFile)) = 0 Then
        strLine = "ECHEC  " & FileNameOf(strFile) & " : fichier introuvable"
        ProcessFile = kxImportFailed
        Exit Function
    End If

    strName = ReadModuleName(strFile)
    If Len(strName) = 0 Then
        strLine = "IGNORE " & FileNameOf(strFile) & " : ligne Attribute VB_Name absente"
        ProcessFile = kxImportSkipped
        Exit Function
    End If
    If StrComp(strName, THIS_MODULE, vbTextCompare) = 0 Then
        strLine = "IGNORE " & strName & " : l'outil ne peut pas se remplacer lui-meme (import manuel)"
        ProcessFile = kxImportSkipped
        Exit Function
    End If

    Set objOld = FindComponent(objProject, strName)
    If Not objOld Is Nothing Then
        If objOld.Type = VBEXT_DOCUMENT Then
            strLine = "IGNORE " & strName & " : module de feuille ou ThisWorkbook (coller le code a la main)"
            ProcessFile = kxImportSkipped
            Exit Function
        End If

        ' 1. Sauvegarde de l'ancien module
        If Len(strBackupFolder) = 0 Then strBackupFolder = CreateBackupFolder()
        strBackupFile = strBackupFolder & Application.PathSeparator & strName & ExtensionOf(objOld.Type)
        objOld.Export strBackupFile

        ' 2. Renommer avant de supprimer : la suppression n'est effective qu'a la
        '    fin de la macro ; sans ce renommage, l'import recevrait le nom "xxx1".
        objOld.Name = TemporaryName(objProject, strName)
        objProject.VBComponents.Remove objOld
        blnRemoved = True
    End If

    ' 3. Import du nouveau module
    Set objNew = objProject.VBComponents.Import(strFile)

    If StrComp(objNew.Name, strName, vbBinaryCompare) <> 0 Then
        strLine = "ECHEC  " & strName & " : importe sous le nom " & objNew.Name & " (a renommer)"
        ProcessFile = kxImportFailed
    ElseIf blnRemoved Then
        strLine = "OK     " & strName & " : remplace"
        ProcessFile = kxImportReplaced
    Else
        strLine = "OK     " & strName & " : ajoute (nouveau module)"
        ProcessFile = kxImportAdded
    End If
    Exit Function

EH:
    strLine = "ECHEC  " & IIf(Len(strName) > 0, strName, FileNameOf(strFile)) & _
              " : erreur " & Err.Number & " - " & Err.Description
    If blnRemoved Then
        strLine = strLine & vbCrLf & "       ATTENTION : l'ancien module a ete supprime. " & _
                  "Le reimporter depuis " & strBackupFile
    End If
    ProcessFile = kxImportFailed
End Function

' Projet VBA du classeur ; Nothing (avec message) si l'acces est impossible
Private Function GetProject() As Object
    Dim objProject As Object

    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox "Enregistrez d'abord le classeur.", vbExclamation, TOOL_TITLE
        Exit Function
    End If

    On Error Resume Next
    Set objProject = ThisWorkbook.VBProject
    If Err.Number <> 0 Or objProject Is Nothing Then
        On Error GoTo 0
        MsgBox "Acces au projet VBA refuse." & vbCrLf & _
               "Cochez 'Acces approuve au modele d'objet du projet VBA' dans le " & _
               "Centre de gestion de la confidentialite, puis relancez.", vbExclamation, TOOL_TITLE
        Exit Function
    End If
    On Error GoTo 0

    If objProject.Protection = VBEXT_PP_LOCKED Then
        MsgBox "Le projet VBA est protege par mot de passe : deverrouillez-le d'abord.", _
               vbExclamation, TOOL_TITLE
        Exit Function
    End If

    Set GetProject = objProject
End Function

' Boite de choix des fichiers (selection multiple). Retourne un tableau de
' chemins, ou Empty si l'utilisateur annule.
Private Function PickModuleFiles() As Variant
    Dim objDialog As Object
    Dim vFiles() As String
    Dim i As Long

    Set objDialog = Application.FileDialog(MSO_FILE_DIALOG_FILE_PICKER)
    With objDialog
        .Title = "Choisir les modules a remplacer ou ajouter"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Modules VBA", "*.bas;*.cls;*.frm"
        If Len(ThisWorkbook.Path) > 0 Then .InitialFileName = ThisWorkbook.Path & Application.PathSeparator
        If .Show <> -1 Then Exit Function

        ReDim vFiles(1 To .SelectedItems.Count)
        For i = 1 To .SelectedItems.Count
            vFiles(i) = .SelectedItems(i)
        Next i
    End With

    PickModuleFiles = vFiles
End Function

' Nom du module lu dans le fichier : Attribute VB_Name = "Nom"
Private Function ReadModuleName(ByVal strFile As String) As String
    Const KEY As String = "Attribute VB_Name = """
    Dim intFile As Integer
    Dim strText As String
    Dim lngStart As Long
    Dim lngEnd As Long

    intFile = FreeFile
    Open strFile For Binary Access Read As #intFile
    strText = Space$(LOF(intFile))
    Get #intFile, , strText
    Close #intFile

    lngStart = InStr(1, strText, KEY, vbTextCompare)
    If lngStart = 0 Then Exit Function
    lngStart = lngStart + Len(KEY)
    lngEnd = InStr(lngStart, strText, """")
    If lngEnd > lngStart Then ReadModuleName = Mid$(strText, lngStart, lngEnd - lngStart)
End Function

' Composant du projet portant ce nom (Nothing s'il n'existe pas)
Private Function FindComponent(ByVal objProject As Object, ByVal strName As String) As Object
    Dim objComp As Object

    For Each objComp In objProject.VBComponents
        If StrComp(objComp.Name, strName, vbTextCompare) = 0 Then
            Set FindComponent = objComp
            Exit Function
        End If
    Next objComp
End Function

' Nom libre pour renommer l'ancien module juste avant sa suppression
Private Function TemporaryName(ByVal objProject As Object, ByVal strName As String) As String
    Dim strCandidate As String
    Dim i As Long

    Do
        i = i + 1
        strCandidate = Left$(strName, MAX_MODULE_NAME - 6) & "_old" & i
    Loop Until FindComponent(objProject, strCandidate) Is Nothing

    TemporaryName = strCandidate
End Function

Private Function CreateBackupFolder() As String
    Dim strFolder As String

    strFolder = ThisWorkbook.Path & Application.PathSeparator & _
                "Backup_VBA_" & Format$(Now, "yyyymmdd_hhnnss")
    If Len(Dir$(strFolder, vbDirectory)) = 0 Then MkDir strFolder
    CreateBackupFolder = strFolder
End Function

Private Function ExtensionOf(ByVal lngType As Long) As String
    Select Case lngType
        Case VBEXT_STD_MODULE:   ExtensionOf = ".bas"
        Case VBEXT_CLASS_MODULE: ExtensionOf = ".cls"
        Case VBEXT_MS_FORM:      ExtensionOf = ".frm"
        Case Else:               ExtensionOf = ".cls"
    End Select
End Function

Private Function FileNameOf(ByVal strPath As String) As String
    FileNameOf = Mid$(strPath, InStrRev(strPath, Application.PathSeparator) + 1)
End Function
