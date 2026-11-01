Attribute VB_Name = "ZmodExportTool"
Option Explicit
Option Private Module

'==============================================================================
' KAMERIX - Outil : export de tous les modules du projet VBA
'------------------------------------------------------------------------------
' Cree un dossier "Export_VBA_aaaammjj_hhmm" a cote du classeur et y exporte :
' modules standard (.bas), classes (.cls), userforms (.frm + .frx) et les
' modules de feuille / ThisWorkbook qui contiennent du code (.cls).
' Prerequis : "Acces approuve au modele d'objet du projet VBA" coche.
'==============================================================================

Private Const VBEXT_STD_MODULE As Long = 1
Private Const VBEXT_CLASS_MODULE As Long = 2
Private Const VBEXT_MS_FORM As Long = 3
Private Const VBEXT_DOCUMENT As Long = 100

Public Sub ExportKamerixModules()
    Dim objProject As Object
    Dim objComp As Object
    Dim strFolder As String
    Dim strExt As String
    Dim lngCount As Long

    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox "Enregistrez d'abord le classeur.", vbExclamation, "Kamerix - export"
        Exit Sub
    End If

    On Error Resume Next
    Set objProject = ThisWorkbook.VBProject
    lngCount = objProject.VBComponents.Count
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "Acces au projet VBA refuse." & vbCrLf & _
               "Cochez 'Acces approuve au modele d'objet du projet VBA' dans le " & _
               "Centre de gestion de la confidentialite, puis relancez.", _
               vbExclamation, "Kamerix - export"
        Exit Sub
    End If
    On Error GoTo EH

    strFolder = ThisWorkbook.Path & Application.PathSeparator & _
                 "Export_VBA_" & Format$(Now, "yyyymmdd_hhnn")
    If Len(Dir$(strFolder, vbDirectory)) = 0 Then MkDir strFolder

    lngCount = 0
    For Each objComp In objProject.VBComponents
        strExt = ""
        Select Case objComp.Type
            Case VBEXT_STD_MODULE:   strExt = ".bas"
            Case VBEXT_CLASS_MODULE: strExt = ".cls"
            Case VBEXT_MS_FORM:      strExt = ".frm"
            Case VBEXT_DOCUMENT
                If objComp.CodeModule.CountOfLines > 1 Then strExt = ".cls"
        End Select

        If Len(strExt) > 0 Then
            objComp.Export strFolder & Application.PathSeparator & objComp.Name & strExt
            lngCount = lngCount + 1
        End If
    Next objComp

    MsgBox lngCount & " module(s) exporte(s) dans :" & vbCrLf & strFolder, _
           vbInformation, "Kamerix - export"
    Exit Sub

EH:
    MsgBox "Echec de l'export (erreur " & Err.Number & ") :" & vbCrLf & Err.Description, _
           vbCritical, "Kamerix - export"
End Sub

