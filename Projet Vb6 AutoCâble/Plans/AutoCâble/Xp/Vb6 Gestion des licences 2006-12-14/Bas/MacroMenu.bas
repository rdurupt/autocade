Attribute VB_Name = "MacroMenu"
Sub SubCreer(MyDroit As String)
boolExec = True

Creer.chargement MyDroit
Unload Creer

boolExec = False
Admin = False
End Sub
Sub Synthese()
boolExec = True

FrmSynthese.Show vbModal


boolExec = False
Admin = False
End Sub
Sub Utilitaire()
boolExec = True

Utilitaires.Show vbModal


boolExec = False
Admin = False
End Sub

Sub SubExportXls()

ExportXls2.Show vbModal
Unload ExportXls2

Admin = False
End Sub
Sub subImportXls()
boolExec = True

ImportXls.Show vbModal
Unload ImportXls

boolExec = False
Admin = False
End Sub
Sub subExporter()
NomenclatureOk = False

boolExec = True

FrmNomChoix.Show vbModal
If FrmNomChoix.Valide = False Then GoTo Fin
ExporterExcel.ChargeNomenclature FrmNomChoix.PreparNomk, ""
Unload ExporterExcel
Fin:
Unload FrmNomChoix


boolExec = False
Admin = False

NomenclatureOk = True
End Sub
Sub subGenerateur()
 boolExec = True

EditeGenerateur.Show vbModal
'SubActionCorrective 505, 0
boolExec = False
Admin = False
End Sub
Sub subGenerateurEtiquette()
 boolExec = True

FrmEtiquette.Show vbModal
Unload FrmEtiquette
'SubActionCorrective 505, 0
boolExec = False
Admin = False
End Sub
Sub subTestWord()
 boolExec = True

CreerFab.Show vbModal
'SubActionCorrective 505, 0
boolExec = False
Admin = False
End Sub
Sub subJob()
Dim FrmLstJob As New LstJob
  boolExec = True

FrmLstJob.Show
'SubActionCorrective 505, 0
boolExec = False
Admin = False
End Sub
Sub subExporterSynthese()
Dim Sql As String
Dim Rs As Recordset
EporteSynthese "SyntG"
Sql = "SELECT Rq_Synthese_Total.* FROM Rq_Synthese_Total;"
Set Rs = Con.OpenRecordSet(Sql)
While Rs.EOF = False
    EporteSynthese "Synt", Rs!Affaire
Rs.MoveNext
Wend
Admin = False
End Sub

Sub subEDITER(MyDroit As String, BooolBloque As Boolean)
Dim MyModifier As Object
Set MyModifier = New Modifier

MyModifier.Visible = False

DoEvents
If BooolBloque = False Then
    CherchPices.Charge MyModifier, "(VerifieDate= Null  and Archiver=false) OR  (IdStatus<4  and Archiver=false)", BooolBloque:=BooolBloque, AvecForm:=True
Else
    CherchPices.Charge MyModifier, "IdStatus<>4", BooolBloque:=BooolBloque, AvecForm:=True
End If
'If BooolBloque = False Then
'CherchPices.Charge MyModifier, "(VerifieDate= Null  and Archiver=false) OR  (IdStatus<4  and Archiver=false)", BooolBloque:=BooolBloque, AvecForm:=True
'Else
'    CherchPices.Charge MyModifier, "IdStatus<>4", BooolBloque:=BooolBloque, AvecForm:=True
'End If
'Set MyFormCible = Nothing
'MyModifier.charger MyDroit

Admin = False
End Sub
Sub subSuperUtilisateur()
Dim MyMenuSuperU As New MenuSuperU
MyMenuSuperU.Show vbModal
End Sub
Sub subUtilisateur()
Dim MyMenuAdmin As New MenuAdmin
MyMenuAdmin.Show vbModal
Admin = False
End Sub
Sub subtestEcart()
MajEcartIndice 341
End Sub
Sub aaaaa()
Shell "E:\Nouveau dossier\Package\setup.exe"
'MDAC_TYP
End Sub
Sub subModifierCartouche()
ModifierCartouches.Show vbModal
End Sub
Sub subVerifierEtude()
Dim MyVerifierEtude As New VerifierEtude
MyVerifierEtude.Show vbModal
Admin = False
End Sub
Sub subImport()
Dim MyImportCablePrixExport As New ImportCablePrixExport
MyImportCablePrixExport.ImporOk = True
MyImportCablePrixExport.Show vbModal
Admin = False
End Sub
Sub subExport()
Dim MyImportCablePrixExport As New ImportCablePrixExport
MyImportCablePrixExport.ImporOk = False
MyImportCablePrixExport.Show vbModal
Admin = False
End Sub

Sub subApprobation()
Dim MyApprobation As New Approbation
MyApprobation.Show vbModal

Admin = False
End Sub
Public Function GestionDesDroit(LibBouton As String) As Boolean
Dim Sql As String
Dim Rs As Recordset
boolExec = True
Sql = "SELECT T_Droits.Id_Bouton "
Sql = Sql & "FROM T_Boutons INNER JOIN T_Droits ON T_Boutons.Id = T_Droits.Id_Bouton "
Sql = Sql & "WHERE T_Boutons.Name='" & MyReplace(LibBouton) & "';"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = False Then
   Useres.charger LibBouton
   GestionDesDroit = Useres.DroitsOk
   Unload Useres
Else
    GestionDesDroit = True
End If
End Function
