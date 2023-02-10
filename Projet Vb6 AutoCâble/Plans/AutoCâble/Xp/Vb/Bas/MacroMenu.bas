Attribute VB_Name = "MacroMenu"
Public Sub SubCreer()
boolExec = True
LoadDb
Useres.Charger "Creer", "Approbateur"
Unload Useres
funCloseDatabase
boolExec = False
Admin = False
End Sub
Public Sub Utilitaire()
boolExec = True
LoadDb
Utilitaires.Show vbModal

funCloseDatabase
boolExec = False
Admin = False
End Sub

Public Sub SubExportXls()
LoadDb
ExportXls2.Show vbModal
Unload ExportXls2
funCloseDatabase
Admin = False
End Sub
Public Sub subImportXls()
boolExec = True
LoadDb
ImportXls.Show vbModal
Unload ImportXls
funCloseDatabase
boolExec = False
Admin = False
End Sub
Public Sub subExporter()
boolExec = True
LoadDb
ExporterExcel.Show vbModal
Unload ExporterExcel
funCloseDatabase
boolExec = False
Admin = False
End Sub
Public Sub subExporterSynthese()
Dim sql As String
Dim rs As Recordset
boolExec = True
LoadDb
EporteSynthese
sql = "SELECT Rq_Synthese_Total.* FROM Rq_Synthese_Total;"
Set rs = Con.OpenRecordSet(sql)
While rs.EOF = False
    EporteSynthese rs!Affaire
rs.MoveNext
Wend
funCloseDatabase
boolExec = False
Admin = False
End Sub

Public Sub subEDITER()
boolExec = True
LoadDb
Modifier.Show vbModal
Unload Modifier
funCloseDatabase
boolExec = False
Admin = False

End Sub
Public Sub subUtilisateur()
boolExec = True
LoadDb
Useres.Charger "MenuAdmin", "Admin"
Unload Useres
funCloseDatabase
boolExec = False
Admin = False
End Sub

Sub aaaaa()
Shell "E:\Nouveau dossier\Package\setup.exe"
'MDAC_TYP
End Sub
Public Sub subModifierCartouche()
boolExec = True
LoadDb
Useres.Charger "ModifierCartouches", "Approbateur"
Unload Useres
funCloseDatabase
boolExec = False
Admin = False
End Sub
Public Sub subVerifierEtude()
boolExec = True
LoadDb
Useres.Charger "Vérificateur", "Vérificateur"
Unload Useres
funCloseDatabase
boolExec = False
Admin = False
End Sub
Public Sub subImport()
boolExec = True
LoadDb
Useres.Charger "ImportCablePrix", "Admin"
Unload Useres
funCloseDatabase
boolExec = False
Admin = False
End Sub
Public Sub subExport()
boolExec = True
LoadDb
Useres.Charger "ExportCablePrix", "Admin"
Unload Useres
funCloseDatabase
boolExec = False
Admin = False
End Sub

Public Sub subApprobation()
boolExec = True
LoadDb
Useres.Charger "Approbation", "Approbateur"
Unload Useres
funCloseDatabase
boolExec = False
Admin = False
End Sub
