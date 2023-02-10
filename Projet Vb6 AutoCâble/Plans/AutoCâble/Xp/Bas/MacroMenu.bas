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
Public Sub SubExportXls()
LoadDb
ExportXls2.Show
Unload ExportXls2
funCloseDatabase
Admin = False
End Sub
Public Sub subImportXls()
boolExec = True
LoadDb
ImportXls.Show
Unload ImportXls
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
Public Sub subApprobation()
boolExec = True
LoadDb
Useres.Charger "Approbation", "Approbateur"
Unload Useres
funCloseDatabase
boolExec = False
Admin = False
End Sub
