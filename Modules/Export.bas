Attribute VB_Name = "Export"
Option Explicit

Sub ExporterModule()
    ' 1. Déclaration des variables :
    Dim vbComp As VBComponent
    Dim cheminExport As String
    Dim typeModule As Long

    ' 2. Parcourt de tous les modules :
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        ' 2.1. Type de module :
        typeModule = vbComp.Type
        ' 2.2. Paramétrage d'exportation :
        Select Case typeModule '// Selon le type de module
            Case vbext_ct_StdModule '// Module standard (.bas)
                ' //Chemin du dossier /modules/ :
                cheminExport = ThisWorkbook.Path & "\Modules\" & vbComp.Name & ".bas"
            Case vbext_ct_ClassModule '// Module de classe (.cls)
                ' //Chemin du dossier /classes/ :
                cheminExport = ThisWorkbook.Path & "\Classes\" & vbComp.Name & ".cls"
            Case vbext_ct_MSForm '// Userform (.frm)
                ' //Chemin du dossier /forms/ :
                cheminExport = ThisWorkbook.Path & "\Forms\" & vbComp.Name & ".frm"
            Case vbext_ct_Document '// Module de document (ThisWorkbook, SHeet1, etc.)
                ' //Chemin du dossier /documents/ :
                cheminExport = ThisWorkbook.Path & "\Documents\" & vbComp.Name & ".txt"
        End Select
        
        ' 2.2. Exportation :
        vbComp.Export cheminExport

    Next vbComp '// vbComp suivant
    
    MsgBox "Modules exportés avec succès."
End Sub
