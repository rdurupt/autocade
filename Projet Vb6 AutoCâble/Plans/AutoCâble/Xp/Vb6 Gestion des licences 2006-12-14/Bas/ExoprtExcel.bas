Attribute VB_Name = "ExoprtExcel"
Global MyExcel As EXCEL.Application
Global MyWorkbook As EXCEL.Workbook
Sub subExporteXls(IdIndiceProjet As Long, Optional NomenclatureOk As Boolean = True)
    Dim Rs As Recordset
    Dim PathPl As String
    Dim Sql As String
    Set TableauPath = funPath
'    Dim ModelAC As String
'    Set TableauPath = funPath
    Set CollectionTor = Nothing
    Set CollectionTor = New Collection
'
     NUMCOM = 0
    NUMNOTA = 0
    NUMNTOR = 0
     NUMNTORBLOC = 0
    PathArchiveAutocad = TableauPath.Item("PathArchiveAutocad")
    PathArchiveAutocad = DefinirChemienComplet(TableauPath.Item("PathServer"), PathArchiveAutocad)
'     If Left(PathArchiveAutocad, 2) <> "\\" And Left(PathArchiveAutocad, 1) <> "\" Then PathArchiveAutocad = TableauPath.Item("PathServer") & PathArchiveAutocad
'     If Right(PathArchiveAutocad, 2) = "\\" Then PathArchiveAutocad = Mid(PathArchiveAutocad, 1, Len(PathArchiveAutocad) - 1)
  
    Sql = "SELECT [T_indiceProjet].[li],[T_indiceProjet].[li] & '_' &  [T_indiceProjet].[LI_Indice] as Liste FROM T_indiceProjet "
    Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"

    
    Set Rs = Con.OpenRecordSet(Sql)
'     ModelAC = TableauPath.Item("ModelAC")
'     ModelAC = DefinirChemienComplet(TableauPath.Item("PathServer"), ModelAC)
'     If Left(ModelAC, 2) <> "\\" And Left(ModelAC, 1) = "\" Then ModelAC = TableauPath.Item("PathServer") & ModelAC
'     If Right(ModelAC, 2) = "\\" Then ModelAC = Mid(ModelAC, 1, Len(ModelAC) - 1)
NbError = 0
If Rs.EOF = True Then Exit Sub
If IsServeur = False Then
    If MsgBox("Voulez-vous exécuter la Macro Exporter Excel" & vbCrLf & Rs!Liste, vbQuestion + vbYesNo, "Auto-Câble") = vbNo Then Exit Sub
End If

Sql = "SELECT RqCartouche.* "
Sql = Sql & "FROM RqCartouche "
 Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = False Then
        PathPl = PathArchive(TableauPath("PathArchiveAutocad"), "" & Rs!Client, "" & Rs!CleAc, "" & Rs!Pieces, "Li", Rs.Fields("Li"), IdIndiceProjet, Rs.Fields("PI_Indice"), Rs.Fields("LI_Indice"), Rs!Version)
   
    ExporteXls PathPl, IdIndiceProjet, PathPl, NomenclatureOk:=NomenclatureOk

    End If
        DoEvents

Fin:
    ReDim TableauDeConnecteurs(0)
    
    
     FormBarGrah.ProgressBar1.Value = 0
     FormBarGrah.ProgressBar1Caption.Caption = " Traitement terminé"
If IsServeur = False Then
MsgBox "Fin du traitement" & vbCrLf & NbError & " Erreur(s) détectée(s) !"
End If
IncrmentServer FormBarGrah, ""
MenuShow = False
End Sub
Public Sub InsertLIgneExcel(MySheet As EXCEL.Worksheet, L As Long)
MySheet.Rows(CStr(L) & ":" & CStr(L)).Insert Shift:=xlDown
End Sub
Function ExporteXls(Xls As String, IdIndiceProjet As Long, Optional PathPl As String, Optional Save As Boolean = True, Optional Edition As Boolean, Optional NomenclatureOk As Boolean = True) As Boolean
Dim Fso As New FileSystemObject
Dim Sql As String
Dim RsIdProjet As Recordset
Dim Rs As Recordset
Dim PathModelXls As String
Dim MySeet As Worksheet
Dim NbEregistrement As Long
Dim RsOnglet As Recordset
Dim MyErr As String
MyPathXlsMoins1 = BackUp(Xls & ".XLS", True, MyPathXlsMoins1)

 Set TableauPath = funPath

 
    Sql = "SELECT T_indiceProjet.Client FROM T_indiceProjet "
    Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"

    NumErr = 1

    Set RsConnecteur = Con.OpenRecordSet(Sql)
    LeCient = UCase(Trim("" & RsConnecteur!Client))
    Sql = "SELECT  T_Clients.PathCatalogue FROM T_Clients "
Sql = Sql & "WHERE T_Clients.Client='" & MyReplace(LeCient) & "';"

       Set RsConnecteur = Con.OpenRecordSet(Sql)
If RsConnecteur.EOF = False Then
    
    If Trim("" & RsConnecteur!PathCatalogue) = "" Then
         DbCatalogue = ""
   Else
             DbCatalogue = RsConnecteur!PathCatalogue
             DbCatalogue = DefinirChemienComplet(TableauPath.Item("PathServer"), DbCatalogue)
'         If Left(DbCatalogue, 2) <> "\\" And Left(DbCatalogue, 1) = "\" Then DbCatalogue = TableauPath.Item("PathServer") & DbCatalogue
'            If Right(DbCatalogue, 2) = "\\" Then DbCatalogue = Mid(DbCatalogue, 1, Len(DbCatalogue) - 1)
    
    End If
Else
    DbCatalogue = ""
End If





PathModelXls = TableauPath.Item("PathModelXls")
PathModelXls = DefinirChemienComplet(TableauPath.Item("PathServer"), PathModelXls)
'         If Left(PathModelXls, 2) <> "\\" And Left(PathModelXls, 1) = "\" Then PathModelXls = TableauPath.Item("PathServer") & PathModelXls
'          If Right(PathModelXls, 2) = "\\" Then PathModelXls = Mid(PathModelXls, 1, Len(PathModelXls) - 1)

'***********************************************************************************************************************
'*                                       Ouvre le Modèle Excel.                                                        *
    If MyPathXlsMoins1 <> "" Then
    If NomenclatureOk = True Then
        Set MyWorkbook = OpenModelXlt(MyPathXlsMoins1)
    Else
        Set MyWorkbook = OpenModelXlt(PathModelXls)
    End If
    Else
        Set MyWorkbook = OpenModelXlt(PathModelXls)
    End If
'    MyWorkbook.Application.Visible = True
'    RetournIdApp "EXCEL.EXE", True
'    MyWorkbook.Application.Visible = True
'***********************************************************************************************************************
'*                                      Exporte la liste des T_Noeuds.                                              *


    Sql = "SELECT T_Noeuds.ACTIVER, T_Noeuds.Fleche_Droite, T_Noeuds.TORON_PRINCIPAL, T_Noeuds.NŒUDS, T_Noeuds.LONGUEUR,  "
    Sql = Sql & "T_Noeuds.LONGUEUR_CUMULEE, T_Noeuds.DESIGN_HAB, T_Noeuds.CODE_RSA, T_Noeuds.CODE_PSA,  "
    Sql = Sql & "T_Noeuds.CODE_ENC, T_Noeuds.DIAMETRE, T_Noeuds.CLASSE_T, T_Noeuds.OPTION,T_Noeuds.Commentaires "
    Sql = Sql & "FROM T_Noeuds "
    Sql = Sql & "WHERE T_Noeuds.Id_IndiceProjet = " & IdIndiceProjet & " "
    Sql = Sql & "ORDER BY T_Noeuds.NŒUDS;"
    Set Rs = Con.OpenRecordSet(Sql)
    
    
    ExporterRecordsetExcel Rs, MyWorkbook, "NOEUDS", IdIndiceProjet

'    ExporteXlsNoeuds Rs, IdIndiceProjet
    
'***********************************************************************************************************************
'*                                      Exporte la liste des Critères.                                              *

    Sql = "SELECT T_Critères.ACTIVER, T_Critères.CODE_CRITERE, T_Critères.CRITERES,T_Critères.DESIGNATION,T_Critères.Commentaires FROM T_Critères "
    Sql = Sql & "WHERE T_Critères.Id_IndiceProjet = " & IdIndiceProjet & " "
    Sql = Sql & "ORDER BY T_Critères.CODE_CRITERE;"
    Set Rs = Con.OpenRecordSet(Sql)

    ExporteXlsCriteres Rs, IdIndiceProjet
    ExporteXlsCriteresFils IdIndiceProjet
'***********************************************************************************************************************
'*                                      Exporte la liste des connecteurs.                                              *

    Sql = "SELECT Connecteurs.ACTIVER, Connecteurs.CONNECTEUR, Connecteurs.RefConnecteurFour,  "
    Sql = Sql & "Connecteurs.[O/N], Connecteurs.DESIGNATION, Connecteurs.CODE_APP, Connecteurs.N°,  "
    Sql = Sql & "Connecteurs.POS, Connecteurs.[POS-OUT], Connecteurs.PRECO1, Connecteurs.PRECO2,  "
    Sql = Sql & "Connecteurs.OPTION, Connecteurs.[100%], Connecteurs.Pylone, Connecteurs.Colonne,  "
    Sql = Sql & "Connecteurs.Ligne, Connecteurs.RefBouchon, Connecteurs.RefBouchonFour, Connecteurs.ReFCapot,  "
    Sql = Sql & "Connecteurs.ReFCapotFour, Connecteurs.RefVerrou, Connecteurs.RefVerrouFour, Connecteurs.LongueurF_Choix,Connecteurs.Commentaires  "
    Sql = Sql & "FROM Connecteurs "
    Sql = Sql & "WHERE Connecteurs.Id_IndiceProjet = " & IdIndiceProjet & " "
    Sql = Sql & "ORDER BY Connecteurs.N°;"
    Set Rs = Con.OpenRecordSet(Sql)
    
    ExporterRecordsetExcel Rs, MyWorkbook, "Connecteurs", IdIndiceProjet

'    ExporteXlsConnecteur Rs, IdIndiceProjet
'***********************************************************************************************************************
'*                                      Exporte la liste des Fils.                                                     *
'
'    Sql = "SELECT Ligne_Tableau_fils.ACTIVER, Ligne_Tableau_fils.LIAI, Ligne_Tableau_fils.DESIGNATION, Ligne_Tableau_fils.FIL,  "
'   Sql = Sql & "Ligne_Tableau_fils.SECT, Ligne_Tableau_fils.TEINT, Ligne_Tableau_fils.TEINT2, Ligne_Tableau_fils.ISO,  "
'   Sql = Sql & "Ligne_Tableau_fils.LONG, Ligne_Tableau_fils.[LONG CP], Ligne_Tableau_fils.COUPE, Ligne_Tableau_fils.POS,  "
'   Sql = Sql & "Ligne_Tableau_fils.[POS-OUT], Ligne_Tableau_fils.FA, Ligne_Tableau_fils.APP, Ligne_Tableau_fils.VOI, Ligne_Tableau_fils.POS2,  "
'   Sql = Sql & "Ligne_Tableau_fils.[POS-OUT2], Ligne_Tableau_fils.FA2, Ligne_Tableau_fils.APP2, Ligne_Tableau_fils.VOI2,  "
'   Sql = Sql & "Ligne_Tableau_fils.PRECO, Ligne_Tableau_fils.OPTION "
'   Sql = Sql & "FROM Ligne_Tableau_fils WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & IdIndiceProjet & " ORDER BY val('' & Ligne_Tableau_fils.FIL);"
'    Set Rs = Con.OpenRecordSet(Sql)
    
'
'     sql = "SELECT Ligne_Tableau_fils.ACTIVER, Ligne_Tableau_fils.LIAI, Ligne_Tableau_fils.DESIGNATION, Ligne_Tableau_fils.FIL,  "
'   sql = sql & "Ligne_Tableau_fils.SECT, Ligne_Tableau_fils.TEINT, Ligne_Tableau_fils.TEINT2, Ligne_Tableau_fils.ISO,  "
'   sql = sql & "Ligne_Tableau_fils.LONG, Ligne_Tableau_fils.[LONG CP], Ligne_Tableau_fils.COUPE, Ligne_Tableau_fils.POS,  "
'   sql = sql & "Ligne_Tableau_fils.[POS-OUT], Ligne_Tableau_fils.FA, Ligne_Tableau_fils.APP, Ligne_Tableau_fils.VOI,Ligne_Tableau_fils. Ligne_Tableau_fils.POS2,  "
'   sql = sql & "Ligne_Tableau_fils.[POS-OUT2], Ligne_Tableau_fils.FA2, Ligne_Tableau_fils.APP2, Ligne_Tableau_fils.VOI2,  "
'   sql = sql & "Ligne_Tableau_fils.PRECO, Ligne_Tableau_fils.OPTION,Ligne_Tableau_fils.[Critères spécifiques] "
'   sql = sql & "FROM Ligne_Tableau_fils WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & IdIndiceProjet & " ORDER BY val('' & Ligne_Tableau_fils.FIL);"
   
   Sql = "SELECT Ligne_Tableau_fils.ACTIVER, Ligne_Tableau_fils.LIAI, Ligne_Tableau_fils.DESIGNATION,  "
   Sql = Sql & "Ligne_Tableau_fils.FIL, Ligne_Tableau_fils.SECT, Ligne_Tableau_fils.TEINT,  "
   Sql = Sql & "Ligne_Tableau_fils.TEINT2, Ligne_Tableau_fils.ISO, Ligne_Tableau_fils.LONG,  "
   Sql = Sql & "Ligne_Tableau_fils.[LONG CP], Ligne_Tableau_fils.Long_Add,Ligne_Tableau_fils.Long_Add2,Ligne_Tableau_fils.COUPE, Ligne_Tableau_fils.POS,  "
   Sql = Sql & "Ligne_Tableau_fils.[POS-OUT], Ligne_Tableau_fils.FA, Ligne_Tableau_fils.APP,Ligne_Tableau_fils.VOI,Ligne_Tableau_fils.[Ref Connecteur], Ligne_Tableau_fils.[Ref Connecteur_Four],  "
   Sql = Sql & "   "
   Sql = Sql & "Ligne_Tableau_fils.[Ref Clip], Ligne_Tableau_fils.[Ref Clip Four],Ligne_Tableau_fils.PRECO,  "
   Sql = Sql & "Ligne_Tableau_fils.[Ref Joint], Ligne_Tableau_fils.[Ref Joint Four],  "
   Sql = Sql & "Ligne_Tableau_fils.POS2, Ligne_Tableau_fils.[POS-OUT2], Ligne_Tableau_fils.FA2,  "
   Sql = Sql & "Ligne_Tableau_fils.APP2,Ligne_Tableau_fils.VOI2, Ligne_Tableau_fils.[Ref Connecteur2], Ligne_Tableau_fils.[Ref Connecteur_Four2],   "
   Sql = Sql & "Ligne_Tableau_fils.[Ref Clip2], Ligne_Tableau_fils.[Ref Clip Four2],Ligne_Tableau_fils.PRECO2,  "
   Sql = Sql & "Ligne_Tableau_fils.[Ref Joint2], Ligne_Tableau_fils.[Ref Joint Four2],  "
   Sql = Sql & "Ligne_Tableau_fils.PRECOG, Ligne_Tableau_fils.OPTION,  "
   Sql = Sql & "Ligne_Tableau_fils.[Critères spécifiques],Ligne_Tableau_fils.Commentaires  "
   Sql = Sql & "FROM Ligne_Tableau_fils "
   Sql = Sql & "Where Ligne_Tableau_fils.Id_IndiceProjet = " & IdIndiceProjet & " "
   Sql = Sql & "ORDER BY Val('' & Ligne_Tableau_fils.FIL);"

    Set Rs = Con.OpenRecordSet(Sql)
    ExporterRecordsetExcel Rs, MyWorkbook, "Ligne_Tableau_fils", IdIndiceProjet
'    ExporteXlsFils Rs, IdIndiceProjet
    
    
    
    
    
'***********************************************************************************************************************
'*                                      Exporte la liste des Composants.                                              *
       
    Sql = "SELECT Composants.ACTIVER, Composants.DESIGNCOMP, Composants.NUMCOMP, Composants.REFCOMP, Composants.OPTION, Composants.Code_APP_Lier, Composants.Voie,Composants.POS, Composants.[POS-OUT],Composants.Commentaires ,Composants.Path "
    Sql = Sql & "FROM Composants "
    Sql = Sql & "WHERE Composants.Id_IndiceProjet = " & IdIndiceProjet & " "
    Sql = Sql & "ORDER BY Composants.NUMCOMP;"
    Set Rs = Con.OpenRecordSet(Sql)
'MyWorkbook.Application.Visible = True
ExporterRecordsetExcel Rs, MyWorkbook, "Composants", IdIndiceProjet


'    ExporteXlsComposants Rs, IdIndiceProjet
    
'***********************************************************************************************************************
'*                                      Exporte la liste des Notas.                                                    *
    
    Sql = "SELECT Nota.ACTIVER, Nota.NOTA, Nota.NUMNOTA ,Nota.OPTION,Nota.Commentaires FROM Nota "
    Sql = Sql & "WHERE Nota.Id_IndiceProjet= " & IdIndiceProjet & " "
    Sql = Sql & "ORDER BY Nota.NUMNOTA ;"

    Set Rs = Con.OpenRecordSet(Sql)
    
    ExporterRecordsetExcel Rs, MyWorkbook, "Notas", IdIndiceProjet
    
'    ExporteXlsNotas Rs, IdIndiceProjet
    
'***********************************************************************************************************************
'*                                      Si NomenclatureOk= faux alors génère la nomenclature.                          *

If NomenclatureOk = False Then
Sql = "SELECT Rq_Cable_Prix.* FROM Rq_Cable_Prix "
    Sql = Sql & "WHERE Rq_Cable_Prix.Id_IndiceProjet=" & IdIndiceProjet & ";"
Set Rs = Con.OpenRecordSet(Sql)
    ExporteXlsPrixFils Rs, IdIndiceProjet

Sql = "SELECT Rq_Habillages_Prix.* FROM Rq_Habillages_Prix "
    Sql = Sql & "Where Rq_Habillages_Prix.Id_IndiceProjet = " & IdIndiceProjet & " "
    Sql = Sql & "ORDER BY Rq_Habillages_Prix.DESIGN_HAB;"
Set Rs = Con.OpenRecordSet(Sql)
    
    ExporteXlsHabillages Rs, IdIndiceProjet
    NomenclatureOk = Nomenclature3(IdIndiceProjet, PathPl, Save)

End If
If Edition = True Then
    Sql = "SELECT T_Nomenclature.CONNECTEUR,T_Nomenclature.[Nb Voies], T_Nomenclature.OPTION, "
    Sql = Sql & "T_Nomenclature.Qté, T_Nomenclature.[Prix U], T_Nomenclature.[Prix Total],  "
    Sql = Sql & "T_Nomenclature.CODE_APP, T_Nomenclature.DESIGNATION, T_Nomenclature.Couleur,  "
    Sql = Sql & "T_Nomenclature.[Lib Connecteur], T_Nomenclature.Fournisseur,  "
    Sql = Sql & "T_Nomenclature.[Ref Four], T_Nomenclature.[Ref Bouch],  "
    Sql = Sql & "T_Nomenclature.[Bouchon Qté], T_Nomenclature.[Bouchon Prix U],  "
    Sql = Sql & "T_Nomenclature.[Bouchon Prix Total], T_Nomenclature.[Lib Bouch],  "
    Sql = Sql & "T_Nomenclature.[Bouch Fourr], T_Nomenclature.[Bouch Réf Four],  "
    Sql = Sql & "T_Nomenclature.[Ref Capot], T_Nomenclature.[Ref Verrou],  "
    Sql = Sql & "T_Nomenclature.[Ref Joint], T_Nomenclature.[Joint Qté],  "
    Sql = Sql & "T_Nomenclature.[Joint Prix U], T_Nomenclature.[Joint Prix Total],  "
    Sql = Sql & "T_Nomenclature.[Lib Joint], T_Nomenclature.[Joint Four],  "
    Sql = Sql & "T_Nomenclature.[Joint Four Réf], T_Nomenclature.[Nb Alvé],  "
    Sql = Sql & "T_Nomenclature.Voie, T_Nomenclature.Famille, T_Nomenclature.[Famille Lib],  "
    Sql = Sql & "T_Nomenclature.[Alvé Réf], T_Nomenclature.[Alvé Qté],  "
    Sql = Sql & "T_Nomenclature.[Alvé Prix U],  "
    Sql = Sql & "T_Nomenclature.[Alvé Prix Total], T_Nomenclature.[Alvé Réf Fourr],  "
    Sql = Sql & "T_Nomenclature.[Alvéole Mini en mm2], T_Nomenclature.[Alvéole Maxi en mm2] "
    Sql = Sql & "FROM T_Nomenclature "
    Sql = Sql & "WHERE T_Nomenclature.Id_IndiceProjet=" & IdIndiceProjet & ";"
    Set Rs = Con.OpenRecordSet(Sql)
    NbEregistrement = 0
    While Rs.EOF = False
        NbEregistrement = NbEregistrement + 1
        Rs.MoveNext
    Wend
    If NbEregistrement = 0 Then NbEregistrement = 1
    FormBarGrah.ProgressBar1.Value = 0
    FormBarGrah.ProgressBar1.Max = NbEregistrement
    FormBarGrah.ProgressBar1Caption.Caption = " Exporter liste Nomenclature Connecteur :"
    Rs.Requery
'    MyWorkbook.Application.Visible = True
    ExporterRecordsetExcel Rs, MyWorkbook, "Nomenclature Connecteur", IdIndiceProjet
    
    ReplaceNull MyWorkbook.Worksheets("Nomenclature Connecteur"), Chr(10), "©"
    
    Sql = "SELECT T_Prix_Fils.TEINT, T_Prix_Fils.OPTION, T_Prix_Fils.ISO, T_Prix_Fils.SECT, T_Prix_Fils.Longeur,  "
    Sql = Sql & "T_Prix_Fils.[Prix U], T_Prix_Fils.[Prix Total] "
    Sql = Sql & "FROM T_Prix_Fils "
    Sql = Sql & "WHERE T_Prix_Fils.Id_IndiceProjet=" & IdIndiceProjet & ";"
    Set Rs = Con.OpenRecordSet(Sql)
    
    NbEregistrement = 0
    While Rs.EOF = False
        NbEregistrement = NbEregistrement + 1
        Rs.MoveNext
    Wend
    If NbEregistrement = 0 Then NbEregistrement = 1
    FormBarGrah.ProgressBar1.Value = 0
    FormBarGrah.ProgressBar1.Max = NbEregistrement
    FormBarGrah.ProgressBar1Caption.Caption = " Exporter liste Nomenclature Fils :"
    Rs.Requery
    
    ExporterRecordsetExcel Rs, MyWorkbook, "Nomenclature Fils", IdIndiceProjet
    ReplaceNull MyWorkbook.Worksheets("Nomenclature Fils"), Chr(10), "©"
    Sql = "SELECT T_Appro_Habillage.DESIGN_HAB, T_Appro_Habillage.OPTION, T_Appro_Habillage.Qté, T_Appro_Habillage.[Prix U],  "
    Sql = Sql & "T_Appro_Habillage.[Prix Total], T_Appro_Habillage.CODE_ENC "
    Sql = Sql & "FROM T_Appro_Habillage "
    Sql = Sql & "WHERE T_Appro_Habillage.Id_IndiceProjet=" & IdIndiceProjet & ";"

    
    Set Rs = Con.OpenRecordSet(Sql)
    
    NbEregistrement = 0
    While Rs.EOF = False
        NbEregistrement = NbEregistrement + 1
        Rs.MoveNext
    Wend
    If NbEregistrement = 0 Then NbEregistrement = 1
    FormBarGrah.ProgressBar1.Value = 0
    FormBarGrah.ProgressBar1.Max = NbEregistrement
    FormBarGrah.ProgressBar1Caption.Caption = " Exporter liste Nomenclature Habillage :"
    Rs.Requery
    
    ExporterRecordsetExcel Rs, MyWorkbook, "Nomenclature Habillage", IdIndiceProjet
ReplaceNull MyWorkbook.Worksheets("Nomenclature Habillage"), Chr(10), "©"


Sql = "SELECT Nomenclature2.LIAI, Nomenclature2.Designation, Nomenclature2.App, Nomenclature2.Voie, Nomenclature2.Ref,  "
    Sql = Sql & "Nomenclature2.RefFour, Nomenclature2.App2, Nomenclature2.Voie2, Nomenclature2.Options, Nomenclature2.ISO,  "
    Sql = Sql & "Nomenclature2.Longueur, Nomenclature2.[Longueur Total], Nomenclature2.TEINT, Nomenclature2.TEINT2,  "
    Sql = Sql & "Nomenclature2.SECT, Nomenclature2.Qts "
    Sql = Sql & "FROM Nomenclature2 "
    Sql = Sql & "WHERE Nomenclature2.Id_IndiceProjet=" & IdIndiceProjet & ";"
     Set Rs = Con.OpenRecordSet(Sql)

    NbEregistrement = 0
    While Rs.EOF = False
        NbEregistrement = NbEregistrement + 1
        Rs.MoveNext
    Wend
    If NbEregistrement = 0 Then NbEregistrement = 1
    FormBarGrah.ProgressBar1.Value = 0
    FormBarGrah.ProgressBar1.Max = NbEregistrement
    FormBarGrah.ProgressBar1Caption.Caption = " Exporter liste Nomenclature :"
    Rs.Requery
        
        ExporterRecordsetExcel Rs, MyWorkbook, "Nomenclature", IdIndiceProjet
'      MyWorkbook.Application.Visible = False
    Sql = "SELECT NomenclaturFinal.Designation, NomenclaturFinal.Famille, NomenclaturFinal.Ref, NomenclaturFinal.RefFour, NomenclaturFinal.Qts,  "
    Sql = Sql & " NomenclaturFinal.ISO, NomenclaturFinal.TEINT, NomenclaturFinal.TEINT2, NomenclaturFinal.SECT,  "
    Sql = Sql & "NomenclaturFinal.Qts_Encelade, NomenclaturFinal.Qts_E_Boutique, NomenclaturFinal.Qts_Appro, NomenclaturFinal.Prix_Revient,  "
    Sql = Sql & "NomenclaturFinal.Prix_Vente, NomenclaturFinal.Options  "
    Sql = Sql & "FROM NomenclaturFinal "
    Sql = Sql & "WHERE NomenclaturFinal.Id_IndiceProjet=" & IdIndiceProjet & ";"
     Set Rs = Con.OpenRecordSet(Sql)
     
    NbEregistrement = 0
    While Rs.EOF = False
        NbEregistrement = NbEregistrement + 1
        Rs.MoveNext
    Wend
    If NbEregistrement = 0 Then NbEregistrement = 1
    FormBarGrah.ProgressBar1.Value = 0
    FormBarGrah.ProgressBar1.Max = NbEregistrement
    FormBarGrah.ProgressBar1Caption.Caption = " Exporter liste Nomenclature Finale :"
    Rs.Requery
    
ExporterRecordsetExcel Rs, MyWorkbook, "Nomenclature Finale", IdIndiceProjet


Sql = "SELECT  T_Dossier_Fabrication.Onglet, T_Dossier_Fabrication.ACTIVER, T_Dossier_Fabrication.LIAI,  "
    Sql = Sql & "T_Dossier_Fabrication.DESIGNATION, T_Dossier_Fabrication.FIL, T_Dossier_Fabrication.SECT, T_Dossier_Fabrication.TEINT,  "
    Sql = Sql & "T_Dossier_Fabrication.TEINT2, T_Dossier_Fabrication.ISO, T_Dossier_Fabrication.LONG, T_Dossier_Fabrication.[LONG CP],  "
    Sql = Sql & "T_Dossier_Fabrication.LONG_ADD, T_Dossier_Fabrication.LONG_ADD2, T_Dossier_Fabrication.COUPE, T_Dossier_Fabrication.POS,  "
    Sql = Sql & "T_Dossier_Fabrication.[POS-OUT], T_Dossier_Fabrication.FA, T_Dossier_Fabrication.APP, T_Dossier_Fabrication.VOI,  "
    Sql = Sql & "T_Dossier_Fabrication.[REF CONNECTEUR], T_Dossier_Fabrication.[REF CONNECTEUR_FOUR], T_Dossier_Fabrication.[REF CLIP],  "
    Sql = Sql & "T_Dossier_Fabrication.[REF CLIP FOUR], T_Dossier_Fabrication.PRECO, T_Dossier_Fabrication.[REF JOINT],  "
    Sql = Sql & "T_Dossier_Fabrication.[REF JOINT FOUR], T_Dossier_Fabrication.POS2, T_Dossier_Fabrication.[POS-OUT2],  "
    Sql = Sql & "T_Dossier_Fabrication.FA2, T_Dossier_Fabrication.APP2, T_Dossier_Fabrication.VOI2, T_Dossier_Fabrication.[REF CONNECTEUR2] ,  "
    Sql = Sql & "T_Dossier_Fabrication.[REF CONNECTEUR_FOUR2], T_Dossier_Fabrication.[REF CLIP2],  "
    Sql = Sql & "T_Dossier_Fabrication.[REF CLIP FOUR2], T_Dossier_Fabrication.PRECO2, T_Dossier_Fabrication.[REF JOINT2],  "
    Sql = Sql & "T_Dossier_Fabrication.[REF JOINT FOUR2], T_Dossier_Fabrication.PRECOG, T_Dossier_Fabrication.Option,  "
    Sql = Sql & "T_Dossier_Fabrication.[CRITÈRES SPÉCIFIQUES] "
    Sql = Sql & "FROM T_Dossier_Fabrication "
    Sql = Sql & "WHERE T_Dossier_Fabrication.Id_IndiceProjet=" & IdIndiceProjet & " "
    Sql = Sql & "ORDER BY T_Dossier_Fabrication.Id;"
    
    
    Sql = "SELECT T_Dossier_Fabrication.Onglet  "
    Sql = Sql & "FROM T_Dossier_Fabrication "
    Sql = Sql & "WHERE T_Dossier_Fabrication.Id_IndiceProjet=" & IdIndiceProjet & "; "

Set Rs = Con.OpenRecordSet(Sql)

    NbEregistrement = 0
    While Rs.EOF = False
        NbEregistrement = NbEregistrement + 1
        Rs.MoveNext
    Wend
    If NbEregistrement = 0 Then NbEregistrement = 1
    FormBarGrah.ProgressBar1.Value = 0
    FormBarGrah.ProgressBar1.Max = NbEregistrement
    FormBarGrah.ProgressBar1Caption.Caption = " Exporter liste Dossier de Fabrication :"
    Rs.Requery
    
    Do While Rs.EOF = False
   IncremanteBarGrah FormBarGrah
   Sql = "SELECT  T_Dossier_Fabrication.Onglet, T_Dossier_Fabrication.ACTIVER, T_Dossier_Fabrication.LIAI,  "
    Sql = Sql & "T_Dossier_Fabrication.DESIGNATION, T_Dossier_Fabrication.FIL, T_Dossier_Fabrication.SECT, T_Dossier_Fabrication.TEINT,  "
    Sql = Sql & "T_Dossier_Fabrication.TEINT2, T_Dossier_Fabrication.ISO, T_Dossier_Fabrication.LONG, T_Dossier_Fabrication.[LONG CP],  "
    Sql = Sql & "T_Dossier_Fabrication.LONG_ADD, T_Dossier_Fabrication.LONG_ADD2, T_Dossier_Fabrication.COUPE, T_Dossier_Fabrication.POS,  "
    Sql = Sql & "T_Dossier_Fabrication.[POS-OUT], T_Dossier_Fabrication.FA, T_Dossier_Fabrication.APP, T_Dossier_Fabrication.VOI,  "
    Sql = Sql & "T_Dossier_Fabrication.[REF CONNECTEUR], T_Dossier_Fabrication.[REF CONNECTEUR_FOUR], T_Dossier_Fabrication.[REF CLIP],  "
    Sql = Sql & "T_Dossier_Fabrication.[REF CLIP FOUR], T_Dossier_Fabrication.PRECO, T_Dossier_Fabrication.[REF JOINT],  "
    Sql = Sql & "T_Dossier_Fabrication.[REF JOINT FOUR], T_Dossier_Fabrication.POS2, T_Dossier_Fabrication.[POS-OUT2],  "
    Sql = Sql & "T_Dossier_Fabrication.FA2, T_Dossier_Fabrication.APP2, T_Dossier_Fabrication.VOI2, T_Dossier_Fabrication.[REF CONNECTEUR2] ,  "
    Sql = Sql & "T_Dossier_Fabrication.[REF CONNECTEUR_FOUR2], T_Dossier_Fabrication.[REF CLIP2],  "
    Sql = Sql & "T_Dossier_Fabrication.[REF CLIP FOUR2], T_Dossier_Fabrication.PRECO2, T_Dossier_Fabrication.[REF JOINT2],  "
    Sql = Sql & "T_Dossier_Fabrication.[REF JOINT FOUR2], T_Dossier_Fabrication.PRECOG, T_Dossier_Fabrication.Option,  "
    Sql = Sql & "T_Dossier_Fabrication.[CRITÈRES SPÉCIFIQUES] "
    Sql = Sql & "FROM T_Dossier_Fabrication "
    Sql = Sql & "WHERE T_Dossier_Fabrication.Id_IndiceProjet=" & IdIndiceProjet & "  and T_Dossier_Fabrication.Onglet ='" & Rs!Onglet & "' "
    Sql = Sql & "ORDER BY T_Dossier_Fabrication.Id;"
  
   
   Set RsOnglet = Con.OpenRecordSet(Sql)
   
'    IncremanteBarGrah FormBarGrah
        ExporterRecordsetExcel RsOnglet, MyWorkbook, Trim("FAB_" & Rs!Onglet), IdIndiceProjet, True, True, "FAB_"
         
        Rs.MoveNext
    Loop
    
    Sql = "SELECT  T_Dossier_Contrôle.Onglet, T_Dossier_Contrôle.ACTIVER, T_Dossier_Contrôle.LIAI,  "
    Sql = Sql & "T_Dossier_Contrôle.DESIGNATION, T_Dossier_Contrôle.FIL, T_Dossier_Contrôle.SECT, T_Dossier_Contrôle.TEINT,  "
    Sql = Sql & "T_Dossier_Contrôle.TEINT2, T_Dossier_Contrôle.ISO, T_Dossier_Contrôle.LONG, T_Dossier_Contrôle.[LONG CP],  "
    Sql = Sql & "T_Dossier_Contrôle.LONG_ADD, T_Dossier_Contrôle.LONG_ADD2, T_Dossier_Contrôle.COUPE, T_Dossier_Contrôle.POS,  "
    Sql = Sql & "T_Dossier_Contrôle.[POS-OUT], T_Dossier_Contrôle.FA, T_Dossier_Contrôle.APP, T_Dossier_Contrôle.VOI,  "
    Sql = Sql & "T_Dossier_Contrôle.[REF CONNECTEUR], T_Dossier_Contrôle.[REF CONNECTEUR_FOUR], T_Dossier_Contrôle.[REF CLIP],  "
    Sql = Sql & "T_Dossier_Contrôle.[REF CLIP FOUR], T_Dossier_Contrôle.PRECO, T_Dossier_Contrôle.[REF JOINT],  "
    Sql = Sql & "T_Dossier_Contrôle.[REF JOINT FOUR], T_Dossier_Contrôle.POS2, T_Dossier_Contrôle.[POS-OUT2],  "
    Sql = Sql & "T_Dossier_Contrôle.FA2, T_Dossier_Contrôle.APP2, T_Dossier_Contrôle.VOI2, T_Dossier_Contrôle.[REF CONNECTEUR2] ,  "
    Sql = Sql & "T_Dossier_Contrôle.[REF CONNECTEUR_FOUR2], T_Dossier_Contrôle.[REF CLIP2],  "
    Sql = Sql & "T_Dossier_Contrôle.[REF CLIP FOUR2], T_Dossier_Contrôle.PRECO2, T_Dossier_Contrôle.[REF JOINT2],  "
    Sql = Sql & "T_Dossier_Contrôle.[REF JOINT FOUR2], T_Dossier_Contrôle.PRECOG, T_Dossier_Contrôle.Option,  "
    Sql = Sql & "T_Dossier_Contrôle.[CRITÈRES SPÉCIFIQUES] "
    Sql = Sql & "FROM T_Dossier_Contrôle "
    Sql = Sql & "WHERE T_Dossier_Contrôle.Id_IndiceProjet=" & IdIndiceProjet & " "
    Sql = Sql & "ORDER BY T_Dossier_Contrôle.Id;"
    
    Sql = "SELECT T_Dossier_Contrôle.Onglet "
    Sql = Sql & "FROM T_Dossier_Contrôle "
    Sql = Sql & "GROUP BY T_Dossier_Contrôle.Onglet, T_Dossier_Contrôle.Id_IndiceProjet "
    Sql = Sql & "HAVING T_Dossier_Contrôle.Id_IndiceProjet=" & IdIndiceProjet & ";"

    
    
    
Set Rs = Con.OpenRecordSet(Sql)

    NbEregistrement = 0
    While Rs.EOF = False
        NbEregistrement = NbEregistrement + 1
        Rs.MoveNext
    Wend
    If NbEregistrement = 0 Then NbEregistrement = 1
    FormBarGrah.ProgressBar1.Value = 0
    FormBarGrah.ProgressBar1.Max = NbEregistrement
    FormBarGrah.ProgressBar1Caption.Caption = " Exporter liste Dossier de Contrôle :"
    
 IncrmentServer FormBarGrah, ""
    Rs.Requery
    
    Do While Rs.EOF = False
    IncremanteBarGrah FormBarGrah
        Sql = "SELECT  T_Dossier_Contrôle.Onglet, T_Dossier_Contrôle.ACTIVER, T_Dossier_Contrôle.LIAI,  "
    Sql = Sql & "T_Dossier_Contrôle.DESIGNATION, T_Dossier_Contrôle.FIL, T_Dossier_Contrôle.SECT, T_Dossier_Contrôle.TEINT,  "
    Sql = Sql & "T_Dossier_Contrôle.TEINT2, T_Dossier_Contrôle.ISO, T_Dossier_Contrôle.LONG, T_Dossier_Contrôle.[LONG CP],  "
    Sql = Sql & "T_Dossier_Contrôle.LONG_ADD, T_Dossier_Contrôle.LONG_ADD2, T_Dossier_Contrôle.COUPE, T_Dossier_Contrôle.POS,  "
    Sql = Sql & "T_Dossier_Contrôle.[POS-OUT], T_Dossier_Contrôle.FA, T_Dossier_Contrôle.APP, T_Dossier_Contrôle.VOI,  "
    Sql = Sql & "T_Dossier_Contrôle.[REF CONNECTEUR], T_Dossier_Contrôle.[REF CONNECTEUR_FOUR], T_Dossier_Contrôle.[REF CLIP],  "
    Sql = Sql & "T_Dossier_Contrôle.[REF CLIP FOUR], T_Dossier_Contrôle.PRECO, T_Dossier_Contrôle.[REF JOINT],  "
    Sql = Sql & "T_Dossier_Contrôle.[REF JOINT FOUR], T_Dossier_Contrôle.POS2, T_Dossier_Contrôle.[POS-OUT2],  "
    Sql = Sql & "T_Dossier_Contrôle.FA2, T_Dossier_Contrôle.APP2, T_Dossier_Contrôle.VOI2, T_Dossier_Contrôle.[REF CONNECTEUR2] ,  "
    Sql = Sql & "T_Dossier_Contrôle.[REF CONNECTEUR_FOUR2], T_Dossier_Contrôle.[REF CLIP2],  "
    Sql = Sql & "T_Dossier_Contrôle.[REF CLIP FOUR2], T_Dossier_Contrôle.PRECO2, T_Dossier_Contrôle.[REF JOINT2],  "
    Sql = Sql & "T_Dossier_Contrôle.[REF JOINT FOUR2], T_Dossier_Contrôle.PRECOG, T_Dossier_Contrôle.Option,  "
    Sql = Sql & "T_Dossier_Contrôle.[CRITÈRES SPÉCIFIQUES] "
    Sql = Sql & "FROM T_Dossier_Contrôle "
    Sql = Sql & "WHERE T_Dossier_Contrôle.Id_IndiceProjet=" & IdIndiceProjet & " and T_Dossier_Contrôle.Onglet ='" & Rs!Onglet & "' "
    Sql = Sql & "ORDER BY T_Dossier_Contrôle.Id;"
   
   Set RsOnglet = Con.OpenRecordSet(Sql)

        ExporterRecordsetExcel RsOnglet, MyWorkbook, Trim("Cont_" & Rs!Onglet), IdIndiceProjet, True, True, "CONT_"
'        Rs.Filter = ""
         If Rs.EOF = True Then Exit Do
        Rs.MoveNext
    Loop
    
End If

Set MySeet = IsertSheet(MyWorkbook, "Prix Fils", True)
'MySeet.Application.Visible = True
MySeet.Delete
Set MySeet = IsertSheet(MyWorkbook, "Appro_Habillage", True)
MySeet.Delete
Set MySeet = IsertSheet(MyWorkbook, "Appro Connectique", True)
MySeet.Delete
Set MySeet = Nothing
'***********************************************************************************************************************
'*                                      Exporte RAPPORT DE_CONTRÔLE_FILAIRE.                                            *
ExporteXlsRAPPORT_DE_CONTROLE_FILAIRE IdIndiceProjet
'***********************************************************************************************************************
'*                                      Exporte Fiche de Contrôle.                                            *
ExporteXlsFiche_de_Controle IdIndiceProjet
On Error Resume Next
'***********************************************************************************************************************
'*                                      Supprime le fichier Excel s'il existe                                          *
If Fso.FileExists(Xls & ".xls") Then Fso.DeleteFile Xls & ".xls"
Set Fso = Nothing

'***********************************************************************************************************************
'*                                      Enregistre le fichier & referme Excel.                                         *
Err.Clear
MyWorkbook.Worksheets(1).Select

MyWorkbook.Application.DisplayAlerts = False
MyWorkbook.SaveAs Xls, ReadOnlyRecommended:=True
If NotSaveRacourci = False Then
 If IdFils <> 0 Then
        Sql = "SELECT RqCartouche.* "
        Sql = Sql & "FROM RqCartouche "
        Sql = Sql & "WHERE T_indiceProjet.Id=" & IdFils & ";"
        Set Rs2 = Con.OpenRecordSet(Sql)
         PathPl2 = PathArchive(PathArchiveAutocad, "" & Rs2!Client, "" & Rs2!CleAc, "" & Rs2!Pieces, "Li", Rs2.Fields("Li"), IdFils, Rs2.Fields("PI_Indice"), Rs2.Fields("LI_Indice"), Rs2!Version)
       Racourci "" & PathPl2, Xls & "", "XLS"
    End If
   
End If
If Err Then
MyErr = Err.Description
     FunError 11, "", MyErr
    If IsServeur = False Then
       MsgBox MyErr
    End If
End If
    
On Error GoTo 0
MyWorkbook.Close False
Set MyWorkbook = Nothing
MyExcel.Quit

Set MyExcel = Nothing
'***********************************************************************************************************************
 FormBarGrah.ProgressBar1.Value = 0
 FormBarGrah.ProgressBar1Caption = " Fin du traitement:"
IncrmentServer FormBarGrah
ExporteXls = True
Sql = "UPDATE T_Job SET T_Job.IdExcel = 0 "
Sql = Sql & "WHERE T_Job.Job=" & Command & ";"
If IsServeur = True Then Con.Execute Sql
End Function
Function ExporteFrmModifier(FRM As Object, IdIndiceProjet As Long, Client As String, FRMAppelant As Object, Optional Edition As Boolean) As Boolean
'Dim Fso As New FileSystemObject
Dim Sql As String
Dim RsIdProjet As Recordset
Dim Rs As Recordset
Dim PathModelXls As String
'Dim 'MySeet As Worksheet
Dim NbEregistrement As Long
Dim RsOnglet As Recordset
Dim MyErr As String
'MyPathXlsMoins1 = BackUp(Xls & ".XLS", True, MyPathXlsMoins1)
'FRM.Visible = False
 Set TableauPath = funPath

 
    Sql = "SELECT T_indiceProjet.Client FROM T_indiceProjet "
    Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"

    NumErr = 1

    Set RsConnecteur = Con.OpenRecordSet(Sql)
    LeCient = UCase(Trim("" & RsConnecteur!Client))
    Sql = "SELECT  T_Clients.PathCatalogue FROM T_Clients "
Sql = Sql & "WHERE T_Clients.Client='" & MyReplace(LeCient) & "';"

       Set RsConnecteur = Con.OpenRecordSet(Sql)
If RsConnecteur.EOF = False Then
    
    If Trim("" & RsConnecteur!PathCatalogue) = "" Then
         DbCatalogue = ""
   Else
             DbCatalogue = RsConnecteur!PathCatalogue
             DbCatalogue = DefinirChemienComplet(TableauPath.Item("PathServer"), DbCatalogue)
'         If Left(DbCatalogue, 2) <> "\\" And Left(DbCatalogue, 1) = "\" Then DbCatalogue = TableauPath.Item("PathServer") & DbCatalogue
'            If Right(DbCatalogue, 2) = "\\" Then DbCatalogue = Mid(DbCatalogue, 1, Len(DbCatalogue) - 1)
    
    End If
Else
    DbCatalogue = ""
End If
Set FRM.CollectionMenu = Nothing
Set FRM.CollectionMenu = New Collection

 FRM.CollectionMenu.Add "Spreadsheet5", "Critères"
 FRM.CollectionMenu.Add "Spreadsheet1", "Connecteurs"
 FRM.CollectionMenu.Add "Spreadsheet2", "Tableau de fils"
 FRM.CollectionMenu.Add "Spreadsheet3", "Composants"
 FRM.CollectionMenu.Add "Spreadsheet4", "Notas"
 FRM.CollectionMenu.Add "Spreadsheet6", "Noeuds"
 FRM.CollectionMenu.Add "Spreadsheet7", "Nomenclature Connecteur"
 FRM.CollectionMenu.Add "Spreadsheet8", "Nomenclature Fils"
 FRM.CollectionMenu.Add "Spreadsheet9", "Nomenclature Habillage"
 FRM.CollectionMenu.Add "Spreadsheet10", "Nomenclatures"
 FRM.CollectionMenu.Add "Spreadsheet11", "Dossier de Fabrication"
 FRM.CollectionMenu.Add "Spreadsheet12", "Dossier de Contrôle"


'
'PathModelXls = TableauPath.Item("PathModelXls")
'PathModelXls = DefinirChemienComplet(TableauPath.Item("PathServer"), PathModelXls)
'         If Left(PathModelXls, 2) <> "\\" And Left(PathModelXls, 1) = "\" Then PathModelXls = TableauPath.Item("PathServer") & PathModelXls
'          If Right(PathModelXls, 2) = "\\" Then PathModelXls = Mid(PathModelXls, 1, Len(PathModelXls) - 1)

'***********************************************************************************************************************
''*                                       Ouvre le Modèle Excel.                                                        *
'    If MyPathXlsMoins1 <> "" Then
'    If NomenclatureOk = True Then
'        Set MyWorkbook = OpenModelXlt(MyPathXlsMoins1)
'    Else
'        Set MyWorkbook = OpenModelXlt(PathModelXls)
'    End If
'    Else
'        Set MyWorkbook = OpenModelXlt(PathModelXls)
'    End If
'    MyWorkbook.Application.Visible = True
'    RetournIdApp "EXCEL.EXE", True
'    MyWorkbook.Application.Visible = True
'***********************************************************************************************************************
'*                                      Exporte la liste des T_Noeuds.                                              *


    Sql = "SELECT T_Noeuds.ACTIVER, T_Noeuds.Fleche_Droite, T_Noeuds.TORON_PRINCIPAL, T_Noeuds.NŒUDS, T_Noeuds.LONGUEUR,  "
    Sql = Sql & "T_Noeuds.LONGUEUR_CUMULEE, T_Noeuds.DESIGN_HAB, T_Noeuds.CODE_RSA, T_Noeuds.CODE_PSA,  "
    Sql = Sql & "T_Noeuds.CODE_ENC, T_Noeuds.DIAMETRE, T_Noeuds.CLASSE_T, T_Noeuds.OPTION,T_Noeuds.Commentaires "
    Sql = Sql & "FROM T_Noeuds "
    Sql = Sql & "WHERE T_Noeuds.Id_IndiceProjet = " & IdIndiceProjet & " "
    Sql = Sql & "ORDER BY T_Noeuds.NŒUDS;"
    Set Rs = Con.OpenRecordSet(Sql)

'    ExporteXlsNoeuds Rs, IdIndiceProjet

    Copy_Rs_Spreadsheet FRM, FRM.Spreadsheet6, Rs, "Noeu", FRMAppelant, "Exportation des nœuds"
    FRM.Spreadsheet6.Columns(FRM.NumCollonne("NoeuActiver")).NumberFormat = "Yes/No"
FRM.Spreadsheet6.Columns(FRM.NumCollonne("NoeuFleche_Droite")).NumberFormat = "Yes/No"
FRM.Spreadsheet6.Columns(FRM.NumCollonne("NoeuTORON_PRINCIPAL")).NumberFormat = "Yes/No"

    
'***********************************************************************************************************************
'*                                      Exporte la liste des Critères.                                              *

    Sql = "SELECT T_Critères.ACTIVER, T_Critères.CODE_CRITERE, T_Critères.CRITERES,T_Critères.DESIGNATION,T_Critères.Commentaires FROM T_Critères "
    Sql = Sql & "WHERE T_Critères.Id_IndiceProjet = " & IdIndiceProjet & " "
    Sql = Sql & "ORDER BY T_Critères.CODE_CRITERE;"
    Set Rs = Con.OpenRecordSet(Sql)
     FRM.Charger_Colection FRM.Spreadsheet1.ActiveSheet, "Con"

    Copy_Rs_Spreadsheet FRM, FRM.Spreadsheet5, Rs, "Crit", FRMAppelant, "Exportation des Critères"
    FRM.Spreadsheet5.Columns(FRM.NumCollonne("CritActiver")).NumberFormat = "Yes/No"
'    ExporteXlsCriteres Rs, IdIndiceProjet
'    ExporteXlsCriteresFils IdIndiceProjet
'***********************************************************************************************************************
'*                                      Exporte la liste des connecteurs.                                              *

    Sql = "SELECT Connecteurs.ACTIVER, Connecteurs.CONNECTEUR, Connecteurs.RefConnecteurFour,  "
    Sql = Sql & "Connecteurs.[O/N], Connecteurs.DESIGNATION, Connecteurs.CODE_APP, Connecteurs.N°,  "
    Sql = Sql & "Connecteurs.POS, Connecteurs.[POS-OUT], Connecteurs.PRECO1, Connecteurs.PRECO2,  "
    Sql = Sql & "Connecteurs.OPTION, Connecteurs.[100%], Connecteurs.Pylone, Connecteurs.Colonne,  "
    Sql = Sql & "Connecteurs.Ligne, Connecteurs.RefBouchon, Connecteurs.RefBouchonFour, Connecteurs.ReFCapot,  "
    Sql = Sql & "Connecteurs.ReFCapotFour, Connecteurs.RefVerrou, Connecteurs.RefVerrouFour, Connecteurs.LongueurF_Choix,Connecteurs.Commentaires  "
    Sql = Sql & "FROM Connecteurs "
    Sql = Sql & "WHERE Connecteurs.Id_IndiceProjet = " & IdIndiceProjet & " "
    Sql = Sql & "ORDER BY Connecteurs.N°;"
    Set Rs = Con.OpenRecordSet(Sql)
    
 
Copy_Rs_Spreadsheet FRM, FRM.Spreadsheet1, Rs, "Con", FRMAppelant, "Exportation des Connecteurs"
FRM.Spreadsheet1.ActiveSheet.Range("a1").Select
FRM.Spreadsheet1.Columns(FRM.NumCollonne("ConActiver")).NumberFormat = "Yes/No"
FRM.Spreadsheet1.Columns(FRM.NumCollonne("ConO/N")).NumberFormat = "Yes/No"
'    ExporteXlsConnecteur Rs, IdIndiceProjet
'***********************************************************************************************************************
'*                                      Exporte la liste des Fils.                                                     *
'
'    Sql = "SELECT Ligne_Tableau_fils.ACTIVER, Ligne_Tableau_fils.LIAI, Ligne_Tableau_fils.DESIGNATION, Ligne_Tableau_fils.FIL,  "
'   Sql = Sql & "Ligne_Tableau_fils.SECT, Ligne_Tableau_fils.TEINT, Ligne_Tableau_fils.TEINT2, Ligne_Tableau_fils.ISO,  "
'   Sql = Sql & "Ligne_Tableau_fils.LONG, Ligne_Tableau_fils.[LONG CP], Ligne_Tableau_fils.COUPE, Ligne_Tableau_fils.POS,  "
'   Sql = Sql & "Ligne_Tableau_fils.[POS-OUT], Ligne_Tableau_fils.FA, Ligne_Tableau_fils.APP, Ligne_Tableau_fils.VOI, Ligne_Tableau_fils.POS2,  "
'   Sql = Sql & "Ligne_Tableau_fils.[POS-OUT2], Ligne_Tableau_fils.FA2, Ligne_Tableau_fils.APP2, Ligne_Tableau_fils.VOI2,  "
'   Sql = Sql & "Ligne_Tableau_fils.PRECO, Ligne_Tableau_fils.OPTION "
'   Sql = Sql & "FROM Ligne_Tableau_fils WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & IdIndiceProjet & " ORDER BY val('' & Ligne_Tableau_fils.FIL);"
'    Set Rs = Con.OpenRecordSet(Sql)
    
'
'     sql = "SELECT Ligne_Tableau_fils.ACTIVER, Ligne_Tableau_fils.LIAI, Ligne_Tableau_fils.DESIGNATION, Ligne_Tableau_fils.FIL,  "
'   sql = sql & "Ligne_Tableau_fils.SECT, Ligne_Tableau_fils.TEINT, Ligne_Tableau_fils.TEINT2, Ligne_Tableau_fils.ISO,  "
'   sql = sql & "Ligne_Tableau_fils.LONG, Ligne_Tableau_fils.[LONG CP], Ligne_Tableau_fils.COUPE, Ligne_Tableau_fils.POS,  "
'   sql = sql & "Ligne_Tableau_fils.[POS-OUT], Ligne_Tableau_fils.FA, Ligne_Tableau_fils.APP, Ligne_Tableau_fils.VOI,Ligne_Tableau_fils. Ligne_Tableau_fils.POS2,  "
'   sql = sql & "Ligne_Tableau_fils.[POS-OUT2], Ligne_Tableau_fils.FA2, Ligne_Tableau_fils.APP2, Ligne_Tableau_fils.VOI2,  "
'   sql = sql & "Ligne_Tableau_fils.PRECO, Ligne_Tableau_fils.OPTION,Ligne_Tableau_fils.[Critères spécifiques] "
'   sql = sql & "FROM Ligne_Tableau_fils WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & IdIndiceProjet & " ORDER BY val('' & Ligne_Tableau_fils.FIL);"
   
   Sql = "SELECT Ligne_Tableau_fils.ACTIVER, Ligne_Tableau_fils.LIAI, Ligne_Tableau_fils.DESIGNATION,  "
   Sql = Sql & "Ligne_Tableau_fils.FIL, Ligne_Tableau_fils.SECT, Ligne_Tableau_fils.TEINT,  "
   Sql = Sql & "Ligne_Tableau_fils.TEINT2, Ligne_Tableau_fils.ISO, Ligne_Tableau_fils.LONG,  "
   Sql = Sql & "Ligne_Tableau_fils.[LONG CP], Ligne_Tableau_fils.Long_Add,Ligne_Tableau_fils.Long_Add2,Ligne_Tableau_fils.COUPE, Ligne_Tableau_fils.POS,  "
   Sql = Sql & "Ligne_Tableau_fils.[POS-OUT], Ligne_Tableau_fils.FA, Ligne_Tableau_fils.APP,Ligne_Tableau_fils.VOI,Ligne_Tableau_fils.[Ref Connecteur], Ligne_Tableau_fils.[Ref Connecteur_Four],  "
   Sql = Sql & "   "
   Sql = Sql & "Ligne_Tableau_fils.[Ref Clip], Ligne_Tableau_fils.[Ref Clip Four],Ligne_Tableau_fils.PRECO,  "
   Sql = Sql & "Ligne_Tableau_fils.[Ref Joint], Ligne_Tableau_fils.[Ref Joint Four],  "
   Sql = Sql & "Ligne_Tableau_fils.POS2, Ligne_Tableau_fils.[POS-OUT2], Ligne_Tableau_fils.FA2,  "
   Sql = Sql & "Ligne_Tableau_fils.APP2,Ligne_Tableau_fils.VOI2, Ligne_Tableau_fils.[Ref Connecteur2], Ligne_Tableau_fils.[Ref Connecteur_Four2],   "
   Sql = Sql & "Ligne_Tableau_fils.[Ref Clip2], Ligne_Tableau_fils.[Ref Clip Four2],Ligne_Tableau_fils.PRECO2,  "
   Sql = Sql & "Ligne_Tableau_fils.[Ref Joint2], Ligne_Tableau_fils.[Ref Joint Four2],  "
   Sql = Sql & "Ligne_Tableau_fils.PRECOG, Ligne_Tableau_fils.OPTION,  "
   Sql = Sql & "Ligne_Tableau_fils.[Critères spécifiques],Ligne_Tableau_fils.Commentaires  "
   Sql = Sql & "FROM Ligne_Tableau_fils "
   Sql = Sql & "Where Ligne_Tableau_fils.Id_IndiceProjet = " & IdIndiceProjet & " "
   Sql = Sql & "ORDER BY Val('' & Ligne_Tableau_fils.FIL);"

    Set Rs = Con.OpenRecordSet(Sql)
'    ExporteXlsFils Rs, IdIndiceProjet

    Copy_Rs_Spreadsheet FRM, FRM.Spreadsheet2, Rs, "Fils", FRMAppelant, "Exportation des Fils"
    FRM.Spreadsheet2.Columns(FRM.NumCollonne("FilsActiver")).NumberFormat = "Yes/No"
    
    
    
'***********************************************************************************************************************
'*                                      Exporte la liste des Composants.                                              *
       
    Sql = "SELECT Composants.ACTIVER, Composants.DESIGNCOMP, Composants.NUMCOMP, Composants.REFCOMP, Composants.OPTION, Composants.Code_APP_Lier, Composants.Voie,Composants.POS, Composants.[POS-OUT],Composants.Commentaires ,Composants.Path "
    Sql = Sql & "FROM Composants "
    Sql = Sql & "WHERE Composants.Id_IndiceProjet = " & IdIndiceProjet & " "
    Sql = Sql & "ORDER BY Composants.NUMCOMP;"
    Set Rs = Con.OpenRecordSet(Sql)
    
 
     Copy_Rs_Spreadsheet FRM, FRM.Spreadsheet3, Rs, "Comp", FRMAppelant, "Exportation des Composants"
    FRM.Spreadsheet3.Columns(FRM.NumCollonne("CompACTIVER")).NumberFormat = "Yes/No"
'    ExporteXlsComposants Rs, IdIndiceProjet
    
'***********************************************************************************************************************
'*                                      Exporte la liste des Notas.                                                    *
    
    Sql = "SELECT Nota.ACTIVER, Nota.NOTA, Nota.NUMNOTA ,Nota.OPTION,Nota.Commentaires FROM Nota "
    Sql = Sql & "WHERE Nota.Id_IndiceProjet= " & IdIndiceProjet & " "
    Sql = Sql & "ORDER BY Nota.NUMNOTA ;"

    Set Rs = Con.OpenRecordSet(Sql)
'    ExporteXlsNotas Rs, IdIndiceProjet

    Copy_Rs_Spreadsheet FRM, FRM.Spreadsheet4, Rs, "Not", FRMAppelant, "Exportation des Notas"
    FRM.Spreadsheet4.Columns(FRM.NumCollonne("NotActiver")).NumberFormat = "Yes/No"
'***********************************************************************************************************************
'*                                      Si NomenclatureOk= faux alors génère la nomenclature.                          *

If NomenclatureOk = False Then
Sql = "SELECT Rq_Cable_Prix.* FROM Rq_Cable_Prix "
    Sql = Sql & "WHERE Rq_Cable_Prix.Id_IndiceProjet=" & IdIndiceProjet & ";"
Set Rs = Con.OpenRecordSet(Sql)
    ExporteXlsPrixFils Rs, IdIndiceProjet

Sql = "SELECT Rq_Habillages_Prix.* FROM Rq_Habillages_Prix "
    Sql = Sql & "Where Rq_Habillages_Prix.Id_IndiceProjet = " & IdIndiceProjet & " "
    Sql = Sql & "ORDER BY Rq_Habillages_Prix.DESIGN_HAB;"
Set Rs = Con.OpenRecordSet(Sql)
    
    ExporteXlsHabillages Rs, IdIndiceProjet
'    NomenclatureOk = Nomenclature3(IdIndiceProjet, PathPl, Save)

End If
If Edition = True Then
    Sql = "SELECT T_Nomenclature.CONNECTEUR,T_Nomenclature.[Nb Voies], T_Nomenclature.OPTION, "
    Sql = Sql & "T_Nomenclature.Qté, T_Nomenclature.[Prix U], T_Nomenclature.[Prix Total],  "
    Sql = Sql & "T_Nomenclature.CODE_APP, T_Nomenclature.DESIGNATION, T_Nomenclature.Couleur,  "
    Sql = Sql & "T_Nomenclature.[Lib Connecteur], T_Nomenclature.Fournisseur,  "
    Sql = Sql & "T_Nomenclature.[Ref Four], T_Nomenclature.[Ref Bouch],  "
    Sql = Sql & "T_Nomenclature.[Bouchon Qté], T_Nomenclature.[Bouchon Prix U],  "
    Sql = Sql & "T_Nomenclature.[Bouchon Prix Total], T_Nomenclature.[Lib Bouch],  "
    Sql = Sql & "T_Nomenclature.[Bouch Fourr], T_Nomenclature.[Bouch Réf Four],  "
    Sql = Sql & "T_Nomenclature.[Ref Capot], T_Nomenclature.[Ref Verrou],  "
    Sql = Sql & "T_Nomenclature.[Ref Joint], T_Nomenclature.[Joint Qté],  "
    Sql = Sql & "T_Nomenclature.[Joint Prix U], T_Nomenclature.[Joint Prix Total],  "
    Sql = Sql & "T_Nomenclature.[Lib Joint], T_Nomenclature.[Joint Four],  "
    Sql = Sql & "T_Nomenclature.[Joint Four Réf], T_Nomenclature.[Nb Alvé],  "
    Sql = Sql & "T_Nomenclature.Voie, T_Nomenclature.Famille, T_Nomenclature.[Famille Lib],  "
    Sql = Sql & "T_Nomenclature.[Alvé Réf], T_Nomenclature.[Alvé Qté],  "
    Sql = Sql & "T_Nomenclature.[Alvé Prix U],  "
    Sql = Sql & "T_Nomenclature.[Alvé Prix Total], T_Nomenclature.[Alvé Réf Fourr],  "
    Sql = Sql & "T_Nomenclature.[Alvéole Mini en mm2], T_Nomenclature.[Alvéole Maxi en mm2] "
    Sql = Sql & "FROM T_Nomenclature "
    Sql = Sql & "WHERE T_Nomenclature.Id_IndiceProjet=" & IdIndiceProjet & ";"
    Set Rs = Con.OpenRecordSet(Sql)
'    NbEregistrement = 0
'    While Rs.EOF = False
'        NbEregistrement = NbEregistrement + 1
'        Rs.MoveNext
'    Wend
'    If NbEregistrement = 0 Then NbEregistrement = 1
'    FormBarGrah.ProgressBar1.Value = 0
'    FormBarGrah.ProgressBar1.Max = NbEregistrement
'    FormBarGrah.ProgressBar1Caption.Caption = " Exporter liste Nomenclature Connecteur :"
'    Rs.Requery
'    MyWorkbook.Application.Visible = True
'Me.Spreadsheet10.Sheets(1).Range("a1").Select
Copy_Rs_Spreadsheet FRM, FRM.Spreadsheet10.Sheets(1), Rs, "NomCon", FRMAppelant, "Nomenclature Connecteur"
'FRM.Charger_Colection FRM.Spreadsheet10.Sheets(1), "NomCon"
'    'ExporterRecordsetExcel Rs, MyWorkbook, "Nomenclature Connecteur", IdIndiceProjet
'
'    'ReplaceNull MyWorkbook.Worksheets("Nomenclature Connecteur"), Chr(10), "©"
    
    Sql = "SELECT T_Prix_Fils.TEINT, T_Prix_Fils.OPTION, T_Prix_Fils.ISO, T_Prix_Fils.SECT, T_Prix_Fils.Longeur,  "
    Sql = Sql & "T_Prix_Fils.[Prix U], T_Prix_Fils.[Prix Total] "
    Sql = Sql & "FROM T_Prix_Fils "
    Sql = Sql & "WHERE T_Prix_Fils.Id_IndiceProjet=" & IdIndiceProjet & ";"
    Set Rs = Con.OpenRecordSet(Sql)
    Copy_Rs_Spreadsheet FRM, FRM.Spreadsheet10.Sheets(2), Rs, "NomFil", FRMAppelant, "Nomenclature Fils"
'    NbEregistrement = 0
'    While Rs.EOF = False
'        NbEregistrement = NbEregistrement + 1
'        Rs.MoveNext
'    Wend
'    If NbEregistrement = 0 Then NbEregistrement = 1
'    FormBarGrah.ProgressBar1.Value = 0
'    FormBarGrah.ProgressBar1.Max = NbEregistrement
'    FormBarGrah.ProgressBar1Caption.Caption = " Exporter liste Nomenclature Fils :"
'    Rs.Requery
    
    'ExporterRecordsetExcel Rs, MyWorkbook, "Nomenclature Fils", IdIndiceProjet
    'ReplaceNull MyWorkbook.Worksheets("Nomenclature Fils"), Chr(10), "©"
    Sql = "SELECT T_Appro_Habillage.DESIGN_HAB, T_Appro_Habillage.OPTION, T_Appro_Habillage.Qté, T_Appro_Habillage.[Prix U],  "
    Sql = Sql & "T_Appro_Habillage.[Prix Total], T_Appro_Habillage.CODE_ENC "
    Sql = Sql & "FROM T_Appro_Habillage "
    Sql = Sql & "WHERE T_Appro_Habillage.Id_IndiceProjet=" & IdIndiceProjet & ";"
Set Rs = Con.OpenRecordSet(Sql)
Copy_Rs_Spreadsheet FRM, FRM.Spreadsheet10.Sheets(3), Rs, "NimHab", FRMAppelant, "Nomenclature Habillage"
'    Copy_Rs_Spreadsheet FRM, FRM.Spreadsheet10(3), Rs, "NimHab"
'
'
'    NbEregistrement = 0
'    While Rs.EOF = False
'        NbEregistrement = NbEregistrement + 1
'        Rs.MoveNext
'    Wend
'    If NbEregistrement = 0 Then NbEregistrement = 1
'    FormBarGrah.ProgressBar1.Value = 0
'    FormBarGrah.ProgressBar1.Max = NbEregistrement
'    FormBarGrah.ProgressBar1Caption.Caption = " Exporter liste Nomenclature Habillage :"
'    Rs.Requery
    
    'ExporterRecordsetExcel Rs, MyWorkbook, "Nomenclature Habillage", IdIndiceProjet
'ReplaceNull MyWorkbook.Worksheets("Nomenclature Habillage"), Chr(10), "©"


Sql = "SELECT Nomenclature2.LIAI, Nomenclature2.Designation, Nomenclature2.App, Nomenclature2.Voie, Nomenclature2.Ref,  "
    Sql = Sql & "Nomenclature2.RefFour, Nomenclature2.App2, Nomenclature2.Voie2, Nomenclature2.Options, Nomenclature2.ISO,  "
    Sql = Sql & "Nomenclature2.Longueur, Nomenclature2.[Longueur Total], Nomenclature2.TEINT, Nomenclature2.TEINT2,  "
    Sql = Sql & "Nomenclature2.SECT, Nomenclature2.Qts "
    Sql = Sql & "FROM Nomenclature2 "
    Sql = Sql & "WHERE Nomenclature2.Id_IndiceProjet=" & IdIndiceProjet & ";"
     Set Rs = Con.OpenRecordSet(Sql)

'    NbEregistrement = 0
'    While Rs.EOF = False
'        NbEregistrement = NbEregistrement + 1
'        Rs.MoveNext
'    Wend
'    If NbEregistrement = 0 Then NbEregistrement = 1
'    FormBarGrah.ProgressBar1.Value = 0
'    FormBarGrah.ProgressBar1.Max = NbEregistrement
'    FormBarGrah.ProgressBar1Caption.Caption = " Exporter liste Nomenclature :"
'    Rs.Requery
Copy_Rs_Spreadsheet FRM, FRM.Spreadsheet10.Sheets(4), Rs, "Nom", FRMAppelant, "Nomenclature"
        'ExporterRecordsetExcel Rs, MyWorkbook, "Nomenclature", IdIndiceProjet
'      MyWorkbook.Application.Visible = False
    Sql = "SELECT NomenclaturFinal.Designation, NomenclaturFinal.Famille, NomenclaturFinal.Ref, NomenclaturFinal.RefFour, NomenclaturFinal.Qts,  "
    Sql = Sql & " NomenclaturFinal.ISO, NomenclaturFinal.TEINT, NomenclaturFinal.TEINT2, NomenclaturFinal.SECT,  "
    Sql = Sql & "NomenclaturFinal.Qts_Encelade, NomenclaturFinal.Qts_E_Boutique, NomenclaturFinal.Qts_Appro, NomenclaturFinal.Prix_Revient,  "
    Sql = Sql & "NomenclaturFinal.Prix_Vente, NomenclaturFinal.Options  "
    Sql = Sql & "FROM NomenclaturFinal "
    Sql = Sql & "WHERE NomenclaturFinal.Id_IndiceProjet=" & IdIndiceProjet & ";"
     Set Rs = Con.OpenRecordSet(Sql)
     Copy_Rs_Spreadsheet FRM, FRM.Spreadsheet10.Sheets(5), Rs, "NomF", FRMAppelant, "Nomenclature Finale"
'    NbEregistrement = 0
'    While Rs.EOF = False
'        NbEregistrement = NbEregistrement + 1
'        Rs.MoveNext
'    Wend
'    If NbEregistrement = 0 Then NbEregistrement = 1
'    FormBarGrah.ProgressBar1.Value = 0
'    FormBarGrah.ProgressBar1.Max = NbEregistrement
'    FormBarGrah.ProgressBar1Caption.Caption = " Exporter liste Nomenclature Finale :"
'    Rs.Requery
    
'ExporterRecordsetExcel Rs, MyWorkbook, "Nomenclature Finale", IdIndiceProjet


Sql = "SELECT  T_Dossier_Fabrication.Onglet, T_Dossier_Fabrication.ACTIVER, T_Dossier_Fabrication.LIAI,  "
    Sql = Sql & "T_Dossier_Fabrication.DESIGNATION, T_Dossier_Fabrication.FIL, T_Dossier_Fabrication.SECT, T_Dossier_Fabrication.TEINT,  "
    Sql = Sql & "T_Dossier_Fabrication.TEINT2, T_Dossier_Fabrication.ISO, T_Dossier_Fabrication.LONG, T_Dossier_Fabrication.[LONG CP],  "
    Sql = Sql & "T_Dossier_Fabrication.LONG_ADD, T_Dossier_Fabrication.LONG_ADD2, T_Dossier_Fabrication.COUPE, T_Dossier_Fabrication.POS,  "
    Sql = Sql & "T_Dossier_Fabrication.[POS-OUT], T_Dossier_Fabrication.FA, T_Dossier_Fabrication.APP, T_Dossier_Fabrication.VOI,  "
    Sql = Sql & "T_Dossier_Fabrication.[REF CONNECTEUR], T_Dossier_Fabrication.[REF CONNECTEUR_FOUR], T_Dossier_Fabrication.[REF CLIP],  "
    Sql = Sql & "T_Dossier_Fabrication.[REF CLIP FOUR], T_Dossier_Fabrication.PRECO, T_Dossier_Fabrication.[REF JOINT],  "
    Sql = Sql & "T_Dossier_Fabrication.[REF JOINT FOUR], T_Dossier_Fabrication.POS2, T_Dossier_Fabrication.[POS-OUT2],  "
    Sql = Sql & "T_Dossier_Fabrication.FA2, T_Dossier_Fabrication.APP2, T_Dossier_Fabrication.VOI2, T_Dossier_Fabrication.[REF CONNECTEUR2] ,  "
    Sql = Sql & "T_Dossier_Fabrication.[REF CONNECTEUR_FOUR2], T_Dossier_Fabrication.[REF CLIP2],  "
    Sql = Sql & "T_Dossier_Fabrication.[REF CLIP FOUR2], T_Dossier_Fabrication.PRECO2, T_Dossier_Fabrication.[REF JOINT2],  "
    Sql = Sql & "T_Dossier_Fabrication.[REF JOINT FOUR2], T_Dossier_Fabrication.PRECOG, T_Dossier_Fabrication.Option,  "
    Sql = Sql & "T_Dossier_Fabrication.[CRITÈRES SPÉCIFIQUES] "
    Sql = Sql & "FROM T_Dossier_Fabrication "
    Sql = Sql & "WHERE T_Dossier_Fabrication.Id_IndiceProjet=" & IdIndiceProjet & " "
    Sql = Sql & "ORDER BY T_Dossier_Fabrication.Id;"
    
    
    Sql = "SELECT T_Dossier_Fabrication.Onglet  "
    Sql = Sql & "FROM T_Dossier_Fabrication "
    Sql = Sql & "WHERE T_Dossier_Fabrication.Id_IndiceProjet=" & IdIndiceProjet & "; "

Set Rs = Con.OpenRecordSet(Sql)

    NbEregistrement = 0
    While Rs.EOF = False
        NbEregistrement = NbEregistrement + 1
        Rs.MoveNext
    Wend
    If NbEregistrement = 0 Then NbEregistrement = 1
    FRMAppelant.ProgressBar1.Value = 0
    FRMAppelant.ProgressBar1.Max = NbEregistrement
    FRMAppelant.ProgressBar1Caption.Caption = " Exporter liste Dossier de Fabrication :"
    Rs.Requery
    
    Do While Rs.EOF = False
   IncremanteBarGrah FormBarGrah
   Sql = "SELECT  T_Dossier_Fabrication.Onglet, T_Dossier_Fabrication.ACTIVER, T_Dossier_Fabrication.LIAI,  "
    Sql = Sql & "T_Dossier_Fabrication.DESIGNATION, T_Dossier_Fabrication.FIL, T_Dossier_Fabrication.SECT, T_Dossier_Fabrication.TEINT,  "
    Sql = Sql & "T_Dossier_Fabrication.TEINT2, T_Dossier_Fabrication.ISO, T_Dossier_Fabrication.LONG, T_Dossier_Fabrication.[LONG CP],  "
    Sql = Sql & "T_Dossier_Fabrication.LONG_ADD, T_Dossier_Fabrication.LONG_ADD2, T_Dossier_Fabrication.COUPE, T_Dossier_Fabrication.POS,  "
    Sql = Sql & "T_Dossier_Fabrication.[POS-OUT], T_Dossier_Fabrication.FA, T_Dossier_Fabrication.APP, T_Dossier_Fabrication.VOI,  "
    Sql = Sql & "T_Dossier_Fabrication.[REF CONNECTEUR], T_Dossier_Fabrication.[REF CONNECTEUR_FOUR], T_Dossier_Fabrication.[REF CLIP],  "
    Sql = Sql & "T_Dossier_Fabrication.[REF CLIP FOUR], T_Dossier_Fabrication.PRECO, T_Dossier_Fabrication.[REF JOINT],  "
    Sql = Sql & "T_Dossier_Fabrication.[REF JOINT FOUR], T_Dossier_Fabrication.POS2, T_Dossier_Fabrication.[POS-OUT2],  "
    Sql = Sql & "T_Dossier_Fabrication.FA2, T_Dossier_Fabrication.APP2, T_Dossier_Fabrication.VOI2, T_Dossier_Fabrication.[REF CONNECTEUR2] ,  "
    Sql = Sql & "T_Dossier_Fabrication.[REF CONNECTEUR_FOUR2], T_Dossier_Fabrication.[REF CLIP2],  "
    Sql = Sql & "T_Dossier_Fabrication.[REF CLIP FOUR2], T_Dossier_Fabrication.PRECO2, T_Dossier_Fabrication.[REF JOINT2],  "
    Sql = Sql & "T_Dossier_Fabrication.[REF JOINT FOUR2], T_Dossier_Fabrication.PRECOG, T_Dossier_Fabrication.Option,  "
    Sql = Sql & "T_Dossier_Fabrication.[CRITÈRES SPÉCIFIQUES] "
    Sql = Sql & "FROM T_Dossier_Fabrication "
    Sql = Sql & "WHERE T_Dossier_Fabrication.Id_IndiceProjet=" & IdIndiceProjet & "  and T_Dossier_Fabrication.Onglet ='" & Rs!Onglet & "' "
    Sql = Sql & "ORDER BY T_Dossier_Fabrication.Id;"
  
   
   Set RsOnglet = Con.OpenRecordSet(Sql)
   
'    IncremanteBarGrah FormBarGrah
        'ExporterRecordsetExcel RsOnglet, MyWorkbook, Trim("FAB_" & Rs!Onglet), IdIndiceProjet, True, True, "FAB_"
         
        Rs.MoveNext
    Loop
    
    Sql = "SELECT  T_Dossier_Contrôle.Onglet, T_Dossier_Contrôle.ACTIVER, T_Dossier_Contrôle.LIAI,  "
    Sql = Sql & "T_Dossier_Contrôle.DESIGNATION, T_Dossier_Contrôle.FIL, T_Dossier_Contrôle.SECT, T_Dossier_Contrôle.TEINT,  "
    Sql = Sql & "T_Dossier_Contrôle.TEINT2, T_Dossier_Contrôle.ISO, T_Dossier_Contrôle.LONG, T_Dossier_Contrôle.[LONG CP],  "
    Sql = Sql & "T_Dossier_Contrôle.LONG_ADD, T_Dossier_Contrôle.LONG_ADD2, T_Dossier_Contrôle.COUPE, T_Dossier_Contrôle.POS,  "
    Sql = Sql & "T_Dossier_Contrôle.[POS-OUT], T_Dossier_Contrôle.FA, T_Dossier_Contrôle.APP, T_Dossier_Contrôle.VOI,  "
    Sql = Sql & "T_Dossier_Contrôle.[REF CONNECTEUR], T_Dossier_Contrôle.[REF CONNECTEUR_FOUR], T_Dossier_Contrôle.[REF CLIP],  "
    Sql = Sql & "T_Dossier_Contrôle.[REF CLIP FOUR], T_Dossier_Contrôle.PRECO, T_Dossier_Contrôle.[REF JOINT],  "
    Sql = Sql & "T_Dossier_Contrôle.[REF JOINT FOUR], T_Dossier_Contrôle.POS2, T_Dossier_Contrôle.[POS-OUT2],  "
    Sql = Sql & "T_Dossier_Contrôle.FA2, T_Dossier_Contrôle.APP2, T_Dossier_Contrôle.VOI2, T_Dossier_Contrôle.[REF CONNECTEUR2] ,  "
    Sql = Sql & "T_Dossier_Contrôle.[REF CONNECTEUR_FOUR2], T_Dossier_Contrôle.[REF CLIP2],  "
    Sql = Sql & "T_Dossier_Contrôle.[REF CLIP FOUR2], T_Dossier_Contrôle.PRECO2, T_Dossier_Contrôle.[REF JOINT2],  "
    Sql = Sql & "T_Dossier_Contrôle.[REF JOINT FOUR2], T_Dossier_Contrôle.PRECOG, T_Dossier_Contrôle.Option,  "
    Sql = Sql & "T_Dossier_Contrôle.[CRITÈRES SPÉCIFIQUES] "
    Sql = Sql & "FROM T_Dossier_Contrôle "
    Sql = Sql & "WHERE T_Dossier_Contrôle.Id_IndiceProjet=" & IdIndiceProjet & " "
    Sql = Sql & "ORDER BY T_Dossier_Contrôle.Id;"
    
    Sql = "SELECT T_Dossier_Contrôle.Onglet "
    Sql = Sql & "FROM T_Dossier_Contrôle "
    Sql = Sql & "GROUP BY T_Dossier_Contrôle.Onglet, T_Dossier_Contrôle.Id_IndiceProjet "
    Sql = Sql & "HAVING T_Dossier_Contrôle.Id_IndiceProjet=" & IdIndiceProjet & ";"

    
    
    
Set Rs = Con.OpenRecordSet(Sql)

    NbEregistrement = 0
    While Rs.EOF = False
        NbEregistrement = NbEregistrement + 1
        Rs.MoveNext
    Wend
    If NbEregistrement = 0 Then NbEregistrement = 1
    FRMAppelant.ProgressBar1.Value = 0
    FRMAppelant.ProgressBar1.Max = NbEregistrement
    FRMAppelant.ProgressBar1Caption.Caption = " Exporter liste Dossier de Contrôle :"
    
 IncrmentServer FormBarGrah, ""
    Rs.Requery
    
    Do While Rs.EOF = False
    IncremanteBarGrah FormBarGrah
        Sql = "SELECT  T_Dossier_Contrôle.Onglet, T_Dossier_Contrôle.ACTIVER, T_Dossier_Contrôle.LIAI,  "
    Sql = Sql & "T_Dossier_Contrôle.DESIGNATION, T_Dossier_Contrôle.FIL, T_Dossier_Contrôle.SECT, T_Dossier_Contrôle.TEINT,  "
    Sql = Sql & "T_Dossier_Contrôle.TEINT2, T_Dossier_Contrôle.ISO, T_Dossier_Contrôle.LONG, T_Dossier_Contrôle.[LONG CP],  "
    Sql = Sql & "T_Dossier_Contrôle.LONG_ADD, T_Dossier_Contrôle.LONG_ADD2, T_Dossier_Contrôle.COUPE, T_Dossier_Contrôle.POS,  "
    Sql = Sql & "T_Dossier_Contrôle.[POS-OUT], T_Dossier_Contrôle.FA, T_Dossier_Contrôle.APP, T_Dossier_Contrôle.VOI,  "
    Sql = Sql & "T_Dossier_Contrôle.[REF CONNECTEUR], T_Dossier_Contrôle.[REF CONNECTEUR_FOUR], T_Dossier_Contrôle.[REF CLIP],  "
    Sql = Sql & "T_Dossier_Contrôle.[REF CLIP FOUR], T_Dossier_Contrôle.PRECO, T_Dossier_Contrôle.[REF JOINT],  "
    Sql = Sql & "T_Dossier_Contrôle.[REF JOINT FOUR], T_Dossier_Contrôle.POS2, T_Dossier_Contrôle.[POS-OUT2],  "
    Sql = Sql & "T_Dossier_Contrôle.FA2, T_Dossier_Contrôle.APP2, T_Dossier_Contrôle.VOI2, T_Dossier_Contrôle.[REF CONNECTEUR2] ,  "
    Sql = Sql & "T_Dossier_Contrôle.[REF CONNECTEUR_FOUR2], T_Dossier_Contrôle.[REF CLIP2],  "
    Sql = Sql & "T_Dossier_Contrôle.[REF CLIP FOUR2], T_Dossier_Contrôle.PRECO2, T_Dossier_Contrôle.[REF JOINT2],  "
    Sql = Sql & "T_Dossier_Contrôle.[REF JOINT FOUR2], T_Dossier_Contrôle.PRECOG, T_Dossier_Contrôle.Option,  "
    Sql = Sql & "T_Dossier_Contrôle.[CRITÈRES SPÉCIFIQUES] "
    Sql = Sql & "FROM T_Dossier_Contrôle "
    Sql = Sql & "WHERE T_Dossier_Contrôle.Id_IndiceProjet=" & IdIndiceProjet & " and T_Dossier_Contrôle.Onglet ='" & Rs!Onglet & "' "
    Sql = Sql & "ORDER BY T_Dossier_Contrôle.Id;"
   
   Set RsOnglet = Con.OpenRecordSet(Sql)

        'ExporterRecordsetExcel RsOnglet, MyWorkbook, Trim("Cont_" & Rs!Onglet), IdIndiceProjet, True, True, "CONT_"
'        Rs.Filter = ""
         If Rs.EOF = True Then Exit Do
        Rs.MoveNext
    Loop
    
End If

'Set 'MySeet = IsertSheet(MyWorkbook, "Prix Fils", True)
''MySeet.Application.Visible = True
'MySeet.Delete
'Set 'MySeet = IsertSheet(MyWorkbook, "Appro_Habillage", True)
'MySeet.Delete
'Set 'MySeet = IsertSheet(MyWorkbook, "Appro Connectique", True)
'MySeet.Delete
'Set 'MySeet = Nothing
'***********************************************************************************************************************
'*                                      Exporte RAPPORT DE_CONTRÔLE_FILAIRE.                                            *
'ExporteXlsRAPPORT_DE_CONTROLE_FILAIRE IdIndiceProjet
'***********************************************************************************************************************
'*                                      Exporte Fiche de Contrôle.                                            *
'ExporteXlsFiche_de_Controle IdIndiceProjet
On Error Resume Next
'***********************************************************************************************************************
'*                                      Supprime le fichier Excel s'il existe                                          *
'If Fso.FileExists(Xls & ".xls") Then Fso.DeleteFile Xls & ".xls"
'Set Fso = Nothing

'***********************************************************************************************************************
'*                                      Enregistre le fichier & referme Excel.                                         *
Err.Clear
'MyWorkbook.Worksheets(1).Select
'
'MyWorkbook.Application.DisplayAlerts = False
'MyWorkbook.SaveAs Xls, ReadOnlyRecommended:=True
'If NotSaveRacourci = False Then
' If IdFils <> 0 Then
'        sql = "SELECT RqCartouche.* "
'        sql = sql & "FROM RqCartouche "
'        sql = sql & "WHERE T_indiceProjet.Id=" & IdFils & ";"
'        Set Rs2 = Con.OpenRecordSet(sql)
'         PathPl2 = PathArchive(PathArchiveAutocad, "" & Rs2!Client, "" & Rs2!CleAc, "" & Rs2!Pieces, "Li", Rs2.Fields("Li"), IdFils, Rs2.Fields("PI_Indice"), Rs2.Fields("LI_Indice"), Rs2!Version)
'       Racourci "" & PathPl2, Xls & "", "XLS"
'    End If
'
'End If
If Err Then
MyErr = Err.Description
     FunError 11, "", MyErr
    If IsServeur = False Then
       MsgBox MyErr
       Resume Next
    End If
End If
    
On Error GoTo 0
'MyWorkbook.Close False
'Set MyWorkbook = Nothing
'MyExcel.Quit
'
'Set MyExcel = Nothing
'***********************************************************************************************************************
 FormBarGrah.ProgressBar1.Value = 0
 FormBarGrah.ProgressBar1Caption = " Fin du traitement:"
IncrmentServer FormBarGrah
NewUserForm2 = True
Sql = "UPDATE T_Job SET T_Job.IdExcel = 0 "
Sql = Sql & "WHERE T_Job.Job=" & Command & ";"
If IsServeur = True Then Con.Execute Sql


 


End Function







Function ExporteXlsPrixFils(Rs As Recordset, Id_IndiceProjet As Long)
Dim MySeet As EXCEL.Worksheet
Dim MyRange As EXCEL.Range
Dim Row As Long
Dim NbLigne As Long
DoEvents
    NbLigne = 0
If Rs.EOF = True Then Exit Function
'While Rs.EOF = False
'NbLigne = NbLigne + 1
'Rs.MoveNext
'Wend

Rs.Requery
'MyWorkbook.Application.Visible = True
Set MySeet = IsertSheet(MyWorkbook, "Prix Fils", True)

DeleteRow MySeet, True

Set MyRange = MySeet.Range("A5").CurrentRegion
'Myrange.Application.Visible = True
For I = 0 To Rs.Fields.Count - 2
    MyRange(1, I + 1) = Rs.Fields(I).Name
Next
Set MyRange = MySeet.Range("A5").CurrentRegion

    MyRange.Interior.ColorIndex = 15
    MyRange.HorizontalAlignment = xlCenter
        
Row = 2
 FormBarGrah.ProgressBar1.Value = 0
 If NbLigne = 0 Then NbLigne = 1
 FormBarGrah.ProgressBar1.Max = NbLigne
 FormBarGrah.ProgressBar1Caption.Caption = " Exporter Prix du Câble :"
While Rs.EOF = False
     IncremanteBarGrah FormBarGrah
    DoEvents
    MyRange(Row, 1) = "" & Rs!TEINT
    MyRange(Row, 2) = "" & Rs!Option
    MyRange(Row, 3) = "" & Rs!ISO
    MyRange(Row, 4) = Val(Replace("" & Rs!SECT, ",", "."))
    MyRange(Row, 5) = Val(Replace("" & Rs!Longeur, ",", "."))
    MyRange(Row, 6) = Val(Replace("" & Rs![Prix u], ",", "."))
    MyRange(Row, 7).FormulaR1C1 = "" & Rs![Prix Total]
    
    Rs.MoveNext
    Row = Row + 1
Wend
Dim Sql As String
Set MyRange = MySeet.Range("A5").CurrentRegion

Sql = "SELECT  T_indiceProjet.Ensemble, T_indiceProjet.Equipement,[Li] & '_' & [LI_Indice] AS Liste, "
Sql = Sql & "[PI] & '_' & [PI_Indice] AS Piece, T_indiceProjet.Client, T_indiceProjet.CleAc "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Id_IndiceProjet & ";"
Dim RsEntetePage As Recordset
Set RsEntetePage = Con.OpenRecordSet(Sql)

MySeet.Range("F2") = "SOUS TOTAL"
FormatExcelPlage MySeet.Range("F2"), 15, False, True, xlCenter, xlCenter
R1 = MySeet.Range(MyRange(2, MyRange.Columns.Count).Address).Row
R2 = MySeet.Range(MyRange(MyRange.Rows.Count, MyRange.Columns.Count).Address).Row
MySeet.Range("G2").FormulaR1C1 = "=SUBTOTAL(9,R[" & R1 - 2 & "]C:R[" & R2 - 2 & "]C)"
FormatExcelPlage MySeet.Range("G2"), 2, False, True, xlCenter, xlCenter

MiseEnPage MySeet, MyRange, "Affaire : " & RsEntetePage!CleAc & Chr(10) & _
 _
     "Pièce : " & RsEntetePage!Piece & Chr(10) & "Liste : " & RsEntetePage!Liste, vbCrLf & "Câblage : " & Replace("" & RsEntetePage!Ensemble, vbCrLf, " ") & Chr(10) & "Equipement : " & Replace("" & RsEntetePage!Equipement, vbCrLf, " ") & Chr(10) & _
     "" _
     , "ENCELADE" & Chr(10) & "Client : " & RsEntetePage!Client & Chr(10) & Format(Date, "dd-mmm-yyyy"), "", "&P/&N", "", 80, "B6", True, 1, True
    
      MaJEncadreXls MyRange, xlThin, xlThin, xlHairline, xlHairline
      
Set RsEntetePage = Con.CloseRecordSet(RsEntetePage)
Set MyRange = Nothing
insertExelAccess MySeet, "T_Prix_Fils", 5, Id_IndiceProjet
Set MySeet = Nothing

End Function
Function ExporteXlsHabillages(Rs As Recordset, Id_IndiceProjet As Long)
Dim MySeet As EXCEL.Worksheet
Dim MyRange As EXCEL.Range
Dim Row As Long
Dim NbLigne As Long
DoEvents
    NbLigne = 0
If Rs.EOF = True Then Exit Function
While Rs.EOF = False
NbLigne = NbLigne + 1
Rs.MoveNext
Wend
Rs.Requery

Set MySeet = IsertSheet(MyWorkbook, "Appro_Habillage", True)
'MySeet.Application.Visible = True
DeleteRow MySeet, True

Set MyRange = MySeet.Range("A1").CurrentRegion
'Myrange.Application.Visible = True
Row = 6
 FormBarGrah.ProgressBar1.Value = 0
 If NbLigne = 0 Then NbLigne = 1
 FormBarGrah.ProgressBar1.Max = NbLigne
 FormBarGrah.ProgressBar1Caption.Caption = " Exporter liste des Habilage :"
 For I = 0 To Rs.Fields.Count - 2
    MySeet.Cells(5, I + 1) = Rs.Fields(I).Name

 Next
While Rs.EOF = False
     IncremanteBarGrah FormBarGrah
     For I = 0 To Rs.Fields.Count - 2
  
     
         If Rs(I).Name = "Prix Total" Then
            MySeet.Cells(Row, I + 1).FormulaR1C1 = "=(RC[-1]*RC[-2])"
         End If
 Next
    DoEvents
    For I = 0 To Rs.Fields.Count - 2
      If Rs(I).Name <> "Prix Total" Then
                                
   
        MySeet.Cells(Row, I + 1) = Trim(Replace("" & Rs(I), vbCrLf, ""))
    End If
    Next
    Row = Row + 1
    Rs.MoveNext
Wend
Dim Sql As String
Set MyRange = MySeet.Range("A5").CurrentRegion

Sql = "SELECT  T_indiceProjet.Ensemble, T_indiceProjet.Equipement,[Li] & '_' & [LI_Indice] AS Liste, "
Sql = Sql & "[PI] & '_' & [PI_Indice] AS Piece, T_indiceProjet.Client, T_indiceProjet.CleAc "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Id_IndiceProjet & ";"
Dim RsEntetePage As Recordset
Set RsEntetePage = Con.OpenRecordSet(Sql)

MySeet.Range("D2") = "SOUS TOTAL"
FormatExcelPlage MySeet.Range("D2"), 15, False, True, xlCenter, xlCenter
R1 = MySeet.Range(MyRange(2, MyRange.Columns.Count).Address).Row
R2 = MySeet.Range(MyRange(MyRange.Rows.Count, MyRange.Columns.Count).Address).Row
MySeet.Range("E2").FormulaR1C1 = "=SUBTOTAL(9,R[" & R1 - 2 & "]C:R[" & R2 - 2 & "]C)"
FormatExcelPlage MySeet.Range("E2"), 2, False, True, xlCenter, xlCenter

MiseEnPage MySeet, MyRange, "Affaire : " & RsEntetePage!CleAc & Chr(10) & _
 _
     "Pièce : " & RsEntetePage!Piece & Chr(10) & "Liste : " & RsEntetePage!Liste, vbCrLf & "Câblage : " & Replace("" & RsEntetePage!Ensemble, vbCrLf, " ") & Chr(10) & "Equipement : " & Replace("" & RsEntetePage!Equipement, vbCrLf, " ") & Chr(10) & _
     "" _
    , "ENCELADE" & Chr(10) & "Client : " & RsEntetePage!Client & Chr(10) & Format(Date, "dd-mmm-yyyy"), "", "&P/&N", "", 80, "B6", True, 1, True
    
      MaJEncadreXls MyRange, xlThin, xlThin, xlHairline, xlHairline
      
Set RsEntetePage = Con.CloseRecordSet(RsEntetePage)
Set MyRange = Nothing
insertExelAccess MySeet, "T_Appro_Habillage", 5, Id_IndiceProjet

Set MySeet = Nothing

    
End Function


Function ExporteXlsFils(Rs As Recordset, Id_IndiceProjet As Long)
Dim MySeet As EXCEL.Worksheet
Dim MyRange As EXCEL.Range
Dim Row As Long
Dim NbLigne As Long
DoEvents
    NbLigne = 0

Set MySeet = IsertSheet(MyWorkbook, "Ligne_Tableau_fils", True)
IsertColonne "ACTIVER", MySeet
DeleteRow MySeet, True

Set MyRange = MySeet.Range("A1").CurrentRegion
ExcelCreatTitre MyRange, Rs
'Myrange.Application.Visible = True
If Rs.EOF = True Then GoTo Fin
'While Rs.EOF = False
'NbLigne = NbLigne + 1
'Rs.MoveNext
'Wend
Rs.Requery

Row = 2
 
 FormBarGrah.ProgressBar1Caption.Caption = " Exporter liste des Fils :"
 FormBarGrah.ProgressBar1.Value = 0
 If NbLigne = 0 Then NbLigne = 1
 FormBarGrah.ProgressBar1.Max = NbLigne
 IncrmentServer FormBarGrah, ""
'While Rs.EOF = False
'     IncremanteBarGrah FormBarGrah
'    IncrmentServer
'    DoEvents
'    ExcelCreatTitre Myrange(Row, 1), Rs, True
''     Myrange(Row, 1) = Replace(Replace("" & Rs!ACTIVER, "Faux", 0), "Vrai", 1)
''     Myrange(Row, 2) = "'" & Rs!Liai
''    Myrange(Row, 3) = "'" & Rs!DESIGNATION
''    Myrange(Row, 4) = "'" & Rs!Fil
''    Myrange(Row, 5) = "'" & Rs!SECT
''    Myrange(Row, 6) = "'" & Rs!TEINT
''    Myrange(Row, 7) = "'" & Rs!TEINT2
''    Myrange(Row, 8) = "'" & Rs!ISO
''    Myrange(Row, 9) = "'" & Rs!Long
''    Myrange(Row, 10) = "'" & Rs![LONG CP]
''    Myrange(Row, 11) = "'" & Rs!Coupe
''    Myrange(Row, 12) = "'" & Rs!POS
''    Myrange(Row, 13) = "'" & Rs![POS-OUT]
''    Myrange(Row, 14) = "'" & Rs!FA
''    Myrange(Row, 15) = "'" & Rs![App]
''    Myrange(Row, 16) = "'" & Rs!VOI
''
''    Myrange(Row, 17) = "'" & Rs![POS2]
''    Myrange(Row, 18) = "'" & Rs![POS-OUT2]
''    Myrange(Row, 19) = "'" & Rs![FA2]
''
''    Myrange(Row, 20) = "'" & Rs![app2]
''    Myrange(Row, 21) = "'" & Rs![VOI2]
''    Myrange(Row, 22) = "'" & Rs![PRECO]
''    Myrange(Row, 23) = "'" & Rs![Option]
'    Rs.MoveNext
'    Row = Row + 1
'Wend
'FormBarGrah.ProgressBar1.Value = 0
' FormBarGrah.ProgressBar1.Max = 1
 If Rs.EOF = False Then

    MySeet.Range("A2").CopyFromRecordset Rs
      ReplaceBool MySeet, "A2:A" & MySeet.Range("A1").CurrentRegion.Rows.Count
End If
'FormBarGrah.ProgressBar1.Value = 1
Dim Sql As String
Fin:
Set MyRange = MySeet.Range("A1").CurrentRegion

Sql = "SELECT  T_indiceProjet.Ensemble, T_indiceProjet.Equipement,[Li] & '_' & [LI_Indice] AS Liste, "
Sql = Sql & "[PI] & '_' & [PI_Indice] AS Piece, T_indiceProjet.Client, T_indiceProjet.CleAc "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Id_IndiceProjet & ";"
Dim RsEntetePage As Recordset
Set RsEntetePage = Con.OpenRecordSet(Sql)

MiseEnPage MySeet, MyRange, "Affaire : " & RsEntetePage!CleAc & Chr(10) & _
 _
     "Pièce : " & RsEntetePage!Piece & Chr(10) & "Liste : " & RsEntetePage!Liste, vbCrLf & "Câblage : " & Replace("" & RsEntetePage!Ensemble, vbCrLf, " ") & Chr(10) & "Equipement : " & Replace("" & RsEntetePage!Equipement, vbCrLf, " ") & Chr(10) & _
      "" _
    , "ENCELADE" & Chr(10) & "Client : " & RsEntetePage!Client & Chr(10) & Format(Date, "dd-mmm-yyyy"), "", "&P/&N", "", 80, "c2", True, 2, True

  MaJEncadreXls MyRange, xlThin, xlThin, xlHairline, xlHairline
  
Set RsEntetePage = Con.CloseRecordSet(RsEntetePage)
Set MyRange = Nothing
Set MySeet = Nothing
End Function
Function ExporteXlsNoeuds(Rs As Recordset, Id_IndiceProjet As Long)

Dim MySeet As EXCEL.Worksheet
Dim MyRange As EXCEL.Range
Dim Row As Long
Dim NbLigne As Long
    NbLigne = 0
  
'MyWorkbook.Application.Visible = True
Set MySeet = IsertSheet(MyWorkbook, "NOEUDS", True)
IsertColonne "ACTIVER", MySeet
DeleteRow MySeet, True
Set MyRange = MySeet.Range("A1").CurrentRegion
ExcelCreatTitre MyRange, Rs
'  If Rs.EOF = True Then GoTo Fin
'While Rs.EOF = False
'NbLigne = NbLigne + 1
'Rs.MoveNext
'Wend
Rs.Requery
'Row = 2
 FormBarGrah.ProgressBar1.Value = 0
 If NbLigne = 0 Then NbLigne = 1
 FormBarGrah.ProgressBar1.Max = NbLigne
 FormBarGrah.ProgressBar1Caption.Caption = " Exporter liste des NOEUDS :"
 IncrmentServer FormBarGrah, ""
'  FormBarGrah.ProgressBar1.Value = 0
'' FormBarGrah.ProgressBar1.Max = 1
 If Rs.EOF = False Then

    MySeet.Range("A2").CopyFromRecordset Rs
    
    ReplaceBool MySeet, "A2:C" & MySeet.Range("A2").CurrentRegion.Rows.Count
        
  
    
End If
'FormBarGrah.ProgressBar1.Value = 1
 
 
 
'While Rs.EOF = False
' IncremanteBarGrah FormBarGrah
' IncrmentServer
'DoEvents
'ExcelCreatTitre Myrange(Row, 1), Rs, True
''    Myrange(Row, 1) = Replace("" & Rs!Fleche_Droite, "Vrai", 1)
''  If Myrange(Row, 1) <> 1 Then Myrange(Row, 1) = 0
''
'' Myrange(Row, 2) = Replace("" & Rs!TORON_PRINCIPAL, "Vrai", 1)
''  If Myrange(Row, 2) <> 1 Then Myrange(Row, 2) = 0
''    Myrange(Row, 3) = Replace("" & Rs!ACTIVER, "Vrai", 1)
''  If Myrange(Row, 3) <> 1 Then Myrange(Row, 3) = 0
''  Myrange(Row, 4) = "'" & Rs!Noeuds
''  Myrange(Row, 5) = Val("" & Rs!Longueur)
''  Myrange(Row, 6) = Val("" & Rs!LONGUEUR_CUMULEE)
''  Myrange(Row, 7) = "" & Rs!DESIGN_HAB
''  Myrange(Row, 8) = "" & Rs!CODE_RSA
''  Myrange(Row, 9) = "" & Rs!CODE_PSA
''  Myrange(Row, 10) = "" & Rs!CODE_ENC
''  Myrange(Row, 11) = "" & Rs!DIAMETRE
''  Myrange(Row, 12) = "" & Rs!CLASSE_T
'    Rs.MoveNext
'    Row = Row + 1
'Wend
Fin:
Set MyRange = MySeet.Range("A1").CurrentRegion
Dim Sql As String
Sql = "SELECT  T_indiceProjet.Ensemble, T_indiceProjet.Equipement,[Li] & '_' & [LI_Indice] AS Liste, "
Sql = Sql & "[PI] & '_' & [PI_Indice] AS Piece, T_indiceProjet.Client, T_indiceProjet.CleAc "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Id_IndiceProjet & ";"
Dim RsEntetePage As Recordset

Set RsEntetePage = Con.OpenRecordSet(Sql)

MiseEnPage MySeet, MyRange, "Affaire : " & RsEntetePage!CleAc & Chr(10) & _
 _
     "Pièce : " & RsEntetePage!Piece & Chr(10) & "Liste : " & RsEntetePage!Liste, vbCrLf & "Câblage : " & Replace("" & RsEntetePage!Ensemble, vbCrLf, " ") & Chr(10) & "Equipement : " & Replace("" & RsEntetePage!Equipement, vbCrLf, " ") & Chr(10) & _
     "" _
    , "ENCELADE" & Chr(10) & "Client : " & RsEntetePage!Client & Chr(10) & Format(Date, "dd-mmm-yyyy"), "", "&P/&N", "", 80, "c2", True, 2, True
    
      MaJEncadreXls MyRange, xlThin, xlThin, xlHairline, xlHairline

Set RsEntetePage = Con.CloseRecordSet(RsEntetePage)
Set MyRange = Nothing
Set MySeet = Nothing
End Function

Sub ExporteXlsCriteresFils(Id_Pieces As Long)
Dim Sql As String
Dim Rs As Recordset
Dim MyRange As Range
Dim MySeet As EXCEL.Worksheet
Dim L As Long
Dim C As Long
Dim Equ As Long
Dim Equ2 As Long
Dim Equipement
Dim Equipement2
Dim Trouve As Boolean
Sql = "SELECT T_indiceProjet.*, T_indiceProjet.Id "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "Where T_indiceProjet.Pere = " & Id_Pieces & "  or T_indiceProjet.Id=" & Id_Pieces & "  "
Sql = Sql & "ORDER BY T_indiceProjet.Id DESC;"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = True Then Exit Sub
Sql = "SELECT  [PI] & '_' & Trim([PL_Indice]) AS Piece "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "Where T_indiceProjet.Id = " & Id_Pieces & " "
Sql = Sql & "Or T_indiceProjet.Pere = " & Id_Pieces & " "
Sql = Sql & "GROUP BY [PI] & '_' & Trim([PL_Indice]) ;"
'Sql = Sql & "ORDER BY T_indiceProjet.Id DESC;"
Set MySeet = IsertSheet(MyWorkbook, "Critères", True)
'MySeet.Application.Visible = True
Set Rs = Con.OpenRecordSet(Sql)
While Rs.EOF = False
    Set MyRange = MySeet.Range("A1").CurrentRegion
    Set MyRange = MyRange(1, MyRange.Columns.Count + 1)
'    Myrange.Application.Visible = True
    ExcelCreatTitre MyRange, Rs, True, True
    Rs.MoveNext
Wend
Sql = "SELECT  T_indiceProjet.Equipement,[PI] & '_' & Trim([PL_Indice]) AS Piece "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "Where T_indiceProjet.Id = " & Id_Pieces & " "
Sql = Sql & "Or T_indiceProjet.Pere = " & Id_Pieces & " "
Sql = Sql & "ORDER BY T_indiceProjet.Id DESC;"
'Myrange.Application.Visible = True
Set MyRange = MySeet.Range("A1").CurrentRegion
Set Rs = Con.OpenRecordSet(Sql)

While Rs.EOF = False
    For C = 1 To MyRange.Columns.Count
        If MyRange(1, C) = "" & Rs!Piece Then Exit For
    Next
    Equipement = Split("" & Rs!Equipement & ";", ";")
    For Equ = 0 To UBound(Equipement) - 1
        Equipement2 = Split("" & Equipement(Equ) & "_", "_")
         
         Set MyRange = MySeet.Range("A1").CurrentRegion
            For L = 2 To MyRange.Rows.Count
            Trouve = False
                If UCase(MyRange(L, 3)) = UCase(Equipement2(0)) Then
                Trouve = True
                If Equipement(Equ) = "" Then Exit For
                    MyRange(L, C) = "X"
                    
             
                    Exit For
                End If
            Next
            If Trouve = False And Equipement(Equ) <> "" Then
            L = MyRange.Rows.Count + 1
                  MyRange(L, C) = "X"
           
                MyRange(L, 1) = 1
                MyRange(L, 2) = UCase(Equipement2(0))
                MyRange(L, 3) = UCase(Equipement2(0))
             
         End If
        
       
    Next
    Rs.MoveNext
Wend
MyRange.AutoFilter
'MyRange.Application.Visible = True
Set MyRange = MySeet.Range("A1").CurrentRegion
Trier MySeet, 1, MyRange.Address, "B1", 1, "", 0, "", 0
Sql = "SELECT  T_indiceProjet.Ensemble, T_indiceProjet.Equipement,[Li] & '_' & [LI_Indice] AS Liste, "
Sql = Sql & "[PI] & '_' & [PI_Indice] AS Piece, T_indiceProjet.Client, T_indiceProjet.CleAc "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Id_Pieces & ";"
Dim RsEntetePage As Recordset

Set RsEntetePage = Con.OpenRecordSet(Sql)

         MyPiedTxt = "Debut : __/__/____" & vbCrLf
MyPiedTxt = MyPiedTxt & "Fin : __/__/____" & vbCrLf
MyPiedTxt = MyPiedTxt & "Réalisé par :"

MiseEnPage MySeet, MyRange, "Affaire : " & RsEntetePage!CleAc & Chr(10) & _
 _
     "Pièce : " & RsEntetePage!Piece & Chr(10) & "Liste : " & RsEntetePage!Liste, vbCrLf & "Câblage : " & Replace("" & RsEntetePage!Ensemble, vbCrLf, " ") & Chr(10) & "Equipement : " & Replace("" & RsEntetePage!Equipement, vbCrLf, " ") & Chr(10) & _
     "" _
    , "ENCELADE" & Chr(10) & "Client : " & RsEntetePage!Client & Chr(10) & Format(Date, "dd-mmm-yyyy"), "", "" & MyPiedTxt, "&P/&N", 100, "A2", True, xlPortrait, True, False, False, 2.5, True, True
    
MaJEncadreXls MyRange, xlThin, xlThin, xlHairline, xlHairline

Set RsEntetePage = Con.CloseRecordSet(RsEntetePage)

End Sub
Sub ExporteFrmCriteresFils(Id_Pieces As Long, MySeet As Object)
Dim Sql As String
Dim Rs As Recordset
Dim MyRange As Object
'Dim MySeet As EXCEL.Worksheet
Dim L As Long
Dim C As Long
Dim Equ As Long
Dim Equ2 As Long
Dim Equipement
Dim Equipement2
Dim Trouve As Boolean
Sql = "SELECT T_indiceProjet.*, T_indiceProjet.Id "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "Where T_indiceProjet.Pere = " & Id_Pieces & "  or T_indiceProjet.Id=" & Id_Pieces & "  "
Sql = Sql & "ORDER BY T_indiceProjet.Id DESC;"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = True Then Exit Sub
Sql = "SELECT  [PI] & '_' & Trim([PL_Indice]) AS Piece "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "Where T_indiceProjet.Id = " & Id_Pieces & " "
Sql = Sql & "Or T_indiceProjet.Pere = " & Id_Pieces & " "
Sql = Sql & "GROUP BY [PI] & '_' & Trim([PL_Indice]) ;"
'Sql = Sql & "ORDER BY T_indiceProjet.Id DESC;"
'Set MySeet = IsertSheet(MyWorkbook, "Critères", True)
'MySeet.Application.Visible = True
Set Rs = Con.OpenRecordSet(Sql)
While Rs.EOF = False
    Set MyRange = MySeet.Range("A1").CurrentRegion
    Set MyRange = MyRange(1, MyRange.Columns.Count + 1)
'    Myrange.Application.Visible = True
    ExcelCreatTitre MyRange, Rs, True, True
    MySeet.Cells(1, MySeet.Range("A1").CurrentRegion.Columns.Count).Interior.Color = ChoixCouleur(0)
    Rs.MoveNext
Wend
Set Rs = Con.CloseRecordSet(Rs)
Sql = "SELECT  T_indiceProjet.Equipement,[PI] & '_' & Trim([PL_Indice]) AS Piece "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "Where T_indiceProjet.Id = " & Id_Pieces & " "
Sql = Sql & "Or T_indiceProjet.Pere = " & Id_Pieces & " "
Sql = Sql & "ORDER BY T_indiceProjet.Id DESC;"
'Myrange.Application.Visible = True
Set MyRange = MySeet.Range("A1").CurrentRegion
Set Rs = Con.OpenRecordSet(Sql)

While Rs.EOF = False
    For C = 1 To MyRange.Columns.Count
        If MyRange(1, C) = "" & Rs!Piece Then Exit For
    Next
    Equipement = Split("" & Rs!Equipement & ";", ";")
    For Equ = 0 To UBound(Equipement) - 1
        Equipement2 = Split("" & Equipement(Equ) & "_", "_")
         
         Set MyRange = MySeet.Range("A1").CurrentRegion
            For L = 2 To MyRange.Rows.Count
            Trouve = False
                If UCase(MyRange(L, 3)) = UCase(Equipement2(0)) Then
                Trouve = True
                If Equipement(Equ) = "" Then Exit For
                    MyRange(L, C) = "X"
                    
             
                    Exit For
                End If
            Next
            If Trouve = False And Equipement(Equ) <> "" Then
            L = MyRange.Rows.Count + 1
                  MyRange(L, C) = "X"
           
                MyRange(L, 1) = 1
                MyRange(L, 2) = UCase(Equipement2(0))
                MyRange(L, 3) = UCase(Equipement2(0))
             
         End If
        
       
    Next
    Rs.MoveNext
Wend

'MyRange.Application.Visible = True
Set MyRange = MySeet.Range("A1").CurrentRegion
MyRange.AutoFilter
MyRange.AutoFilter
On Error Resume Next
    MyRange.CurrentRegion.AutoFitColumns
    Err.Clear
    MyRange.Cells.EntireColumn.AutoFit
    Err.Clear
    On Error GoTo 0

Set MyRange = Nothing
'Trier MySeet, 1, Myrange.Address, "B1", 1, "", 0, "", 0
'sql = "SELECT  T_indiceProjet.Ensemble, T_indiceProjet.Equipement,[Li] & '_' & [LI_Indice] AS Liste, "
'sql = sql & "[PI] & '_' & [PI_Indice] AS Piece, T_indiceProjet.Client, T_indiceProjet.CleAc "
'sql = sql & "FROM T_indiceProjet "
'sql = sql & "WHERE T_indiceProjet.Id=" & Id_Pieces & ";"
'Dim RsEntetePage As Recordset
'
'Set RsEntetePage = Con.OpenRecordSet(sql)
'
'         MyPiedTxt = "Debut : __/__/____" & vbCrLf
'MyPiedTxt = MyPiedTxt & "Fin : __/__/____" & vbCrLf
'MyPiedTxt = MyPiedTxt & "Réalisé par :"
'
'MiseEnPage MySeet, Myrange, "Affaire : " & RsEntetePage!CleAc & Chr(10) & _
' _
'     "Pièce : " & RsEntetePage!Piece & Chr(10) & "Liste : " & RsEntetePage!Liste, vbCrLf & "Câblage : " & Replace("" & RsEntetePage!Ensemble, vbCrLf, " ") & Chr(10) & "Equipement : " & Replace("" & RsEntetePage!Equipement, vbCrLf, " ") & Chr(10) & _
'     "" _
'    , "ENCELADE" & Chr(10) & "Client : " & RsEntetePage!Client & Chr(10) & Format(Date, "dd-mmm-yyyy"), "", "" & MyPiedTxt, "&P/&N", 100, "A2", True, xlPortrait, True, False, False, 2.5, True, True
'
'MaJEncadreXls Myrange, xlThin, xlThin, xlHairline, xlHairline
'
'Set RsEntetePage = Con.CloseRecordSet(RsEntetePage)

End Sub

Function ExporteXlsCriteres(Rs As Recordset, Id_IndiceProjet As Long)
Dim MySeet As EXCEL.Worksheet
Dim MyRange As EXCEL.Range
Dim Row As Long
Dim NbLigne As Long
    NbLigne = 0
  
'MyWorkbook.Application.Visible = True
Set MySeet = IsertSheet(MyWorkbook, "Critères", True)

DeleteRow MySeet, True

Set MyRange = MySeet.Range("A1").CurrentRegion
ExcelCreatTitre MyRange, Rs
IsertColonne "ACTIVER", MySeet
'MySeet.Application.Visible = True
'InsertColonneApres MySeet, "ACTIVER", "toto"
  If Rs.EOF = True Then GoTo Fin
While Rs.EOF = False
NbLigne = NbLigne + 1
Rs.MoveNext
Wend
Rs.Requery
Row = 2
 FormBarGrah.ProgressBar1.Value = 0
 If NbLigne = 0 Then NbLigne = 1
 FormBarGrah.ProgressBar1.Max = NbLigne
 FormBarGrah.ProgressBar1Caption.Caption = " Exporter liste des Critères :"
 FormBarGrah.ProgressBar1.Value = 0
IncrmentServer FormBarGrah, ""
' FormBarGrah.ProgressBar1.Max = 1
' If Rs.EOF = False Then
'
'    MySeet.Range("A2").CopyFromRecordset Rs
'End If
'FormBarGrah.ProgressBar1.Value = 1
While Rs.EOF = False
 IncremanteBarGrah FormBarGrah
IncrmentServer FormBarGrah, ""
DoEvents
ExcelCreatTitre MyRange(Row, 1), Rs, True

'     Myrange(Row, 1) = Replace(Replace("" & Rs!ACTIVER, "Faux", 0), "Vrai", 1)
'  Myrange(Row, 2) = "'" & Rs!CODE_CRITERE
'  Myrange(Row, 3) = "'" & Rs!CRITERES

    Rs.MoveNext
    Row = Row + 1
Wend
Fin:
Set MyRange = MySeet.Range("A1").CurrentRegion
Dim Sql As String
Sql = "SELECT  T_indiceProjet.Ensemble, T_indiceProjet.Equipement,[Li] & '_' & [LI_Indice] AS Liste, "
Sql = Sql & "[PI] & '_' & [PI_Indice] AS Piece, T_indiceProjet.Client, T_indiceProjet.CleAc "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Id_IndiceProjet & ";"
Dim RsEntetePage As Recordset

Set RsEntetePage = Con.OpenRecordSet(Sql)

MiseEnPage MySeet, MyRange, "Affaire : " & RsEntetePage!CleAc & Chr(10) & _
 _
     "Pièce : " & RsEntetePage!Piece & Chr(10) & "Liste : " & RsEntetePage!Liste, vbCrLf & "Câblage : " & Replace("" & RsEntetePage!Ensemble, vbCrLf, " ") & Chr(10) & "Equipement : " & Replace("" & RsEntetePage!Equipement, vbCrLf, " ") & Chr(10) & _
     "" _
    , "ENCELADE" & Chr(10) & "Client : " & RsEntetePage!Client & Chr(10) & Format(Date, "dd-mmm-yyyy"), "", "&P/&N", "", 80, "C2", True, 2, True

MaJEncadreXls MyRange, xlThin, xlThin, xlHairline, xlHairline

Set RsEntetePage = Con.CloseRecordSet(RsEntetePage)
Set MyRange = Nothing
Set MySeet = Nothing

End Function
Sub IsertColonne(NameCol As String, sheet As Worksheet, Optional Fin As Boolean, Optional Celles As String)
Dim MyRange As Range
Dim Trouve As Boolean
Dim MyCell As String
Trouve = False
Set MyRange = sheet.Range("A1").CurrentRegion
For I = 1 To MyRange.Columns.Count
    If UCase(MyRange(1, I)) = UCase(NameCol) Then
        Trouve = True
        Exit For
    End If
Next I
If Trouve = False Then
    If Trim(Celles) <> "" Then
        MyCell = Celles
    Else
    If Fin = False Then
        MyCell = "A1"
    Else
        MyCell = MyRange(1, MyRange.Columns.Count + 1).Address
        
    End If
    End If
'    If Sheet.Range(MyCell).AutoFilter = True Then Sheet.Range(MyCell).AutoFilter
   
    MyCell = Replace(MyCell, "$", "")
    sheet.Range(MyCell).Insert Shift:=xlToRight
    sheet.Range(MyCell) = NameCol
    sheet.Range(MyCell).Interior.ColorIndex = 15
    sheet.Range(MyCell).HorizontalAlignment = xlCenter
    
    
'    If Sheet.Range(MyCell).AutoFilter = False Then Sheet.Range(MyCell).AutoFilter
    
End If
    
End Sub
Function ExporteXlsRAPPORT_DE_CONTROLE_FILAIRE(Id_IndiceProjet As Long)
Dim MySeet As EXCEL.Worksheet
Dim MyRange As EXCEL.Range
Dim Row As Long
Set MySeet = IsertSheet(MyWorkbook, "RAPPORT DE CONTRÔLE CONTINUITE", True)
Set MyRange = MySeet.Range("A1").CurrentRegion
'MySeet.Application.Visible = True
'Myrange.Application.Visible = True
Dim Sql As String
Sql = "SELECT  T_indiceProjet.Ensemble, T_indiceProjet.Equipement,[Li] & '_' & [LI_Indice] AS Liste, "
Sql = Sql & "[PI] & '_' & [PI_Indice] AS Piece, T_indiceProjet.Client, T_indiceProjet.CleAc "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Id_IndiceProjet & ";"
Dim RsEntetePage As Recordset

Set RsEntetePage = Con.OpenRecordSet(Sql)

MiseEnPage MySeet, MyRange, "Affaire : " & RsEntetePage!CleAc & Chr(10) & _
 _
     "Pièce : " & RsEntetePage!Piece & Chr(10) & "Liste : " & RsEntetePage!Liste, vbCrLf & "Câblage : " & Replace("" & RsEntetePage!Ensemble, vbCrLf, " ") & Chr(10) & "Equipement : " & Replace("" & RsEntetePage!Equipement, vbCrLf, " ") & Chr(10) & _
     "" _
    , "ENCELADE" & Chr(10) & "Client : " & RsEntetePage!Client & Chr(10) & Format(Date, "dd-mmm-yyyy"), "", "&P/&N", "", 80, "A2", True, 2, False, False, False, 2.5, False, False
    
MaJEncadreXls MyRange, xlThin, xlThin, xlHairline, xlHairline

Set RsEntetePage = Con.CloseRecordSet(RsEntetePage)
Set MyRange = Nothing
Set MySeet = Nothing
End Function
Function ExporteXlsFiche_de_Controle(Id_IndiceProjet As Long)
Dim MySeet As EXCEL.Worksheet
Dim MyRange As EXCEL.Range
Dim Row As Long
Set MySeet = IsertSheet(MyWorkbook, "Fiche de Contrôle", True)
Set MyRange = MySeet.Range("A1").CurrentRegion

Set MyRange = MySeet.Range("A1").CurrentRegion
Dim Sql As String
Sql = "SELECT  T_indiceProjet.Ensemble, T_indiceProjet.Equipement,[Li] & '_' & [LI_Indice] AS Liste, "
Sql = Sql & "[PI] & '_' & [PI_Indice] AS Piece, T_indiceProjet.Client, T_indiceProjet.CleAc "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Id_IndiceProjet & ";"
Dim RsEntetePage As Recordset

Set RsEntetePage = Con.OpenRecordSet(Sql)

         MyPiedTxt = "Debut : __/__/____" & vbCrLf
MyPiedTxt = MyPiedTxt & "Fin : __/__/____" & vbCrLf
MyPiedTxt = MyPiedTxt & "Réalisé par :"

MiseEnPage MySeet, MyRange, "Affaire : " & RsEntetePage!CleAc & Chr(10) & _
 _
     "Pièce : " & RsEntetePage!Piece & Chr(10) & "Liste : " & RsEntetePage!Liste, vbCrLf & "Câblage : " & Replace("" & RsEntetePage!Ensemble, vbCrLf, " ") & Chr(10) & "Equipement : " & Replace("" & RsEntetePage!Equipement, vbCrLf, " ") & Chr(10) & _
     "" _
    , "ENCELADE" & Chr(10) & "Client : " & RsEntetePage!Client & Chr(10) & Format(Date, "dd-mmm-yyyy"), "", "" & MyPiedTxt, "&P/&N", 100, "A2", True, xlPortrait, False, False, False, 2.5, False, False
    
'MaJEncadreXls Myrange, xlThin, xlThin, xlHairline, xlHairline

Set RsEntetePage = Con.CloseRecordSet(RsEntetePage)
Set MyRange = Nothing
Set MySeet = Nothing
End Function


Function ExporteXlsConnecteur(Rs As Recordset, Id_IndiceProjet As Long)
Dim MySeet As EXCEL.Worksheet
Dim MyRange As EXCEL.Range
Dim Row As Long
Dim NbLigne As Long
    NbLigne = 0
    
Set MySeet = IsertSheet(MyWorkbook, "Connecteurs", True)
IsertColonne "ACTIVER", MySeet
DeleteRow MySeet, True

Set MyRange = MySeet.Range("A1").CurrentRegion
MyRange.Application.Visible = False
ExcelCreatTitre MyRange, Rs
'If Rs.EOF = True Then GoTo Fin
'While Rs.EOF = False
'NbLigne = NbLigne + 1
'Rs.MoveNext
'Wend
Rs.Requery
Row = 2
 FormBarGrah.ProgressBar1.Value = 0
 If NbLigne = 0 Then NbLigne = 1
 FormBarGrah.ProgressBar1.Max = NbLigne
 FormBarGrah.ProgressBar1Caption.Caption = " Exporter liste des Connecteurs :"
 IncrmentServer FormBarGrah, ""
'While Rs.EOF = False
' IncremanteBarGrah FormBarGrah
' IncrmentServer
'DoEvents
''Myrange.Application.Visible = True
'ExcelCreatTitre Myrange(Row, 1), Rs, True
'' Myrange(Row, 1) = Replace(Replace("" & Rs!ACTIVER, "Faux", 0), "Vrai", 1)
''  Myrange(Row, 2) = "'" & Rs!Connecteur
''  Myrange(Row, 3) = Abs(Rs![O/N])
''  Myrange(Row, 4) = "'" & Rs!DESIGNATION
''  Myrange(Row, 5) = "'" & Rs!Code_APP
''  Myrange(Row, 6) = "'" & Rs![N°]
''  Myrange(Row, 7) = "'" & Rs!POS
''  Myrange(Row, 8) = "'" & Rs![POS-OUT]
''  Myrange(Row, 9) = "'" & Rs!PRECO1
''  Myrange(Row, 10) = "'" & Rs!PRECO2
''  Myrange(Row, 11) = "'" & Rs![Option]
''  Myrange(Row, 12) = "'" & Rs![100%]
''  Myrange(Row, 13) = "'" & Rs![Pylone]
''  Myrange(Row, 14) = "'" & Rs![Colonne]
'' Myrange(Row, 15) = "'" & Rs![Ligne]
'    Rs.MoveNext
'    Row = Row + 1
'Wend
' FormBarGrah.ProgressBar1.Value = 0
' FormBarGrah.ProgressBar1.Max = 1
 If Rs.EOF = False Then

    MySeet.Range("A2").CopyFromRecordset Rs
'    MySeet.Application.Visible = True
   
    ReplaceBool MySeet, "A2:A" & MySeet.Range("A1").CurrentRegion.Rows.Count & ",d2:d" & MySeet.Range("A1").CurrentRegion.Rows.Count
   
        
   
End If
'FormBarGrah.ProgressBar1.Value = 1
Fin:
Set MyRange = MySeet.Range("A1").CurrentRegion
Dim Sql As String
Sql = "SELECT  T_indiceProjet.Ensemble, T_indiceProjet.Equipement,[Li] & '_' & [LI_Indice] AS Liste, "
Sql = Sql & "[PI] & '_' & [PI_Indice] AS Piece, T_indiceProjet.Client, T_indiceProjet.CleAc "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Id_IndiceProjet & ";"
Dim RsEntetePage As Recordset

Set RsEntetePage = Con.OpenRecordSet(Sql)
Dim TxtPieCentre As String
TxtPieCentre = "N° de Série Câblage : " & Chr(10)
TxtPieCentre = TxtPieCentre & "Date Début Travaux :" & Chr(10)
TxtPieCentre = TxtPieCentre & "Date Fin Travaux :" & Chr(10)
TxtPieCentre = TxtPieCentre & "Technicien(s) :" & Chr(10)
'TxtPieCentre = TxtPieCentre & "&[Page]/&[Pages]"


MiseEnPage MySeet, MyRange, "Affaire : " & RsEntetePage!CleAc & Chr(10) & _
 _
     "Pièce : " & RsEntetePage!Piece & Chr(10) & "Liste : " & RsEntetePage!Liste, vbCrLf & "Câblage : " & Replace("" & RsEntetePage!Ensemble, vbCrLf, " ") & Chr(10) & "Equipement : " & Replace("" & RsEntetePage!Equipement, vbCrLf, " ") & Chr(10) & _
     "" _
    , "ENCELADE" & Chr(10) & "Client : " & RsEntetePage!Client & Chr(10) & Format(Date, "dd-mmm-yyyy"), "", "" & TxtPieCentre, "&P/&N", 80, "C2", True, 2, True, , , 2.5

MaJEncadreXls MyRange, xlThin, xlThin, xlHairline, xlHairline

Set RsEntetePage = Con.CloseRecordSet(RsEntetePage)
Set MyRange = Nothing
Set MySeet = Nothing
End Function
Function ExporteXlsNotas(Rs As Recordset, Id_IndiceProjet As Long)
Dim MySeet As EXCEL.Worksheet
Dim MyRange As EXCEL.Range
Dim Row As Long
Dim NbLigne As Long
    NbLigne = 0
   
Set MySeet = IsertSheet(MyWorkbook, "Notas", True)
DeleteRow MySeet, True
IsertColonne "ACTIVER", MySeet


Set MyRange = MySeet.Range("A1").CurrentRegion
ExcelCreatTitre MyRange, Rs
 If Rs.EOF = True Then GoTo Fin
'While Rs.EOF = False
'NbLigne = NbLigne + 1
'Rs.MoveNext
'Wend
'Rs.Requery

If Rs.EOF = False Then
FormBarGrah.ProgressBar1Caption.Caption = " Exporter liste des Notas :"
 FormBarGrah.ProgressBar1.Value = 0
 If NbLigne = 0 Then NbLigne = 1
 FormBarGrah.ProgressBar1.Max = NbLigne
 IncrmentServer FormBarGrah, ""
    MySeet.Range("A2").CopyFromRecordset Rs
     ReplaceBool MySeet, "A2:A" & MySeet.Range("A1").CurrentRegion.Rows.Count
End If
Row = 2
' FormBarGrah.ProgressBar1.Value = 0
' If NbLigne = 0 Then NbLigne = 1
' FormBarGrah.ProgressBar1.Max = NbLigne
' FormBarGrah.ProgressBar1Caption.Caption = " Exporter liste des Connecteurs :"
' FormBarGrah.ProgressBar1.Value = 0
' FormBarGrah.ProgressBar1.Max = 1
' If Rs.EOF = False Then
'
'    MySeet.Range("A2").CopyFromRecordset Rs
'End If
'FormBarGrah.ProgressBar1.Value = 1
'While Rs.EOF = False
' IncremanteBarGrah FormBarGrah
' IncrmentServer
'DoEvents
'ExcelCreatTitre Myrange(Row, 1), Rs, True
''     Myrange(Row, 1) = Replace(Replace("" & Rs!ACTIVER, "Faux", 0), "Vrai", 1)
''  Myrange(Row, 2) = "'" & Rs!Nota
''  Myrange(Row, 3) = "'" & Rs!NUMNOTA
'
'    Rs.MoveNext
'    Row = Row + 1
'Wend
Fin:
Set MyRange = MySeet.Range("A1").CurrentRegion
Dim Sql As String
Sql = "SELECT  T_indiceProjet.Ensemble, T_indiceProjet.Equipement,[Li] & '_' & [LI_Indice] AS Liste, "
Sql = Sql & "[PI] & '_' & [PI_Indice] AS Piece, T_indiceProjet.Client, T_indiceProjet.CleAc "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Id_IndiceProjet & ";"
Dim RsEntetePage As Recordset
Set RsEntetePage = Con.OpenRecordSet(Sql)

MiseEnPage MySeet, MyRange, "Affaire : " & RsEntetePage!CleAc & Chr(10) & _
 _
     "Pièce : " & RsEntetePage!Piece & Chr(10) & "Liste : " & RsEntetePage!Liste, vbCrLf & "Câblage : " & Replace("" & RsEntetePage!Ensemble, vbCrLf, " ") & Chr(10) & "Equipement : " & Replace("" & RsEntetePage!Equipement, vbCrLf, " ") & Chr(10) & _
     "" _
    , "ENCELADE" & Chr(10) & "Client : " & RsEntetePage!Client & Chr(10) & Format(Date, "dd-mmm-yyyy"), "", "&P/&N", "", 80, "C2", True, 2

MaJEncadreXls MyRange, xlThin, xlThin, xlHairline, xlHairline

Set RsEntetePage = Con.CloseRecordSet(RsEntetePage)
Set MyRange = Nothing
Set MySeet = Nothing
End Function
Public Sub ExporterRecordsetExcel(Rs As Recordset, MyWorbook As Workbook, Onglet As String, Id_IndiceProjet As Long, _
                                    Optional SurOnglet As Boolean, Optional AficherVueArriere As Boolean, Optional Prefix As String)
Dim MySheet As Worksheet
Dim MyRange As Range

Set MySheet = IsertSheet(MyWorkbook, Onglet, True)
Set MyRange = MySheet.Range("A1").CurrentRegion
'MySheet.Application.Visible = True
DeleteRow MySheet, True
Set MyRange = MySheet.Range("A1").CurrentRegion
ExcelCreatTitre MyRange, Rs, False, True, Formule:=True
'If SurOnglet = True Then
'    Rs.Filter = "Onglet=" & Onglet
'End If

Const sDelimiteur$ = vbTab


MySheet.Range("A2").CopyFromRecordset Rs
'Do While Rs.EOF = False
' IncremanteBarGrah FormBarGrah
'    If SurOnglet = True Then
'        If UCase("FAB_" & Rs!Onglet) <> UCase(Trim(Onglet)) Then
'            If UCase("CONT_" & Rs!Onglet) <> UCase(Trim(Onglet)) Then Exit Do
'        End If
'    End If
'    Set Myrange = MySheet.Range("A1").CurrentRegion
'    Set Myrange = MySheet.Range(Myrange(Myrange.Rows.Count + 1, 1).Address & ":" & Myrange(Myrange.Rows.Count + 1, Myrange.Columns.Count).Address)
'
'    ExcelCreatTitre Myrange, Rs, True, Formule:=True
'    Rs.MoveNext
'Loop
 Set MyRange = MySheet.Range("A1").CurrentRegion
For I = 0 To Rs.Fields.Count - 1
    If Rs(I).Type = adBoolean Then
        MySheet.Columns(I + 1).Replace "VRAI", 1
        MySheet.Columns(I + 1).Replace "FAUX", 0
        MySheet.Columns(I + 1).Replace "YES", 1
        MySheet.Columns(I + 1).Replace "NO", 0
        MySheet.Columns(I + 1).Replace "TRUE", 1
        MySheet.Columns(I + 1).Replace "FALSE", 0
        
    End If
    
Next

'  MySheet.Cells.FormulaR1C1 = MySheet.Cells.Value
Dim C_Serch As Long
C_Serch = 1
While C_Serch <> 0

 C_Serch = SerchXlsColumn(MyRange, MyRange(1, 1), "=(RC[-1]*RC[-2])")
 If C_Serch <> 0 Then
    MySheet.Range(Replace(MyRange(2, C_Serch).Address, "$", "") & ":" & Replace(MyRange(MyRange.Rows.Count, C_Serch).Address, "$", "")).FormulaR1C1 = "=(RC[-1]*RC[-2])"
 End If
 Wend
   
Dim Sql As String
Sql = "SELECT  T_indiceProjet.Ensemble, T_indiceProjet.Equipement,[Li] & '_' & [LI_Indice] AS Liste, "
Sql = Sql & "[PI] & '_' & [PI_Indice] AS Piece, T_indiceProjet.Client, T_indiceProjet.CleAc "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Id_IndiceProjet & ";"
Dim RsEntetePage As Recordset
Set RsEntetePage = Con.OpenRecordSet(Sql)
If AficherVueArriere = True Then
    VueArriere MySheet
End If
Set MyRange = MySheet.Range("A1").CurrentRegion
If PortraitPaysage = 0 Then PortraitPaysage = 2
MiseEnPage MySheet, MyRange, "Affaire : " & RsEntetePage!CleAc & Chr(10) & _
 _
     "Pièce : " & RsEntetePage!Piece & Chr(10) & "Liste : " & RsEntetePage!Liste, vbCrLf & "Câblage : " & Replace("" & RsEntetePage!Ensemble, vbCrLf, " ") & Chr(10) & "Equipement : " & Replace("" & RsEntetePage!Equipement, vbCrLf, " ") & Chr(10) & _
     "" _
    , "ENCELADE" & Chr(10) & "Client : " & RsEntetePage!Client & Chr(10) & Format(Date, "dd-mmm-yyyy"), "", "&P/&N", "", 80, "C2", True, PortraitPaysage = 2, True, False

MaJEncadreXls MyRange, xlThin, xlThin, xlHairline, xlHairline

Set RsEntetePage = Con.CloseRecordSet(RsEntetePage)
Set MyRange = Nothing
Set MySeet = Nothing
End Sub
Function ExporteXlsComposants(Rs As Recordset, Id_IndiceProjet As Long)
Dim MySeet As EXCEL.Worksheet
Dim MyRange As EXCEL.Range
Dim Row As Long
Dim Rep As String
Dim NbColonne As Long
Dim TxtPoin As String
Dim NbLigne As Long
Dim Sql As String
Dim Fso As New FileSystemObject
Dim PathComposantsDefault As String
Dim RsComposants As Recordset
Dim C As Long
    NbLigne = 0
 Set MySeet = IsertSheet(MyWorkbook, "Composants", True)
' MySeet.Application.Visible = True

' IsertColonne "ACTIVER", MySeet
' IsertColonne "OPTION", MySeet, True
 
DeleteRow MySeet, True
'MySeet.Application.Visible = True
ExcelCreatTitre MySeet.Range("A1").CurrentRegion, Rs
'DeletCol MySeet, "K"
Set MyRange = MySeet.Range("A1").CurrentRegion
MyRange.Interior.ColorIndex = 15
'Myrange.Application.Visible = True
NbColonne = MyRange.Columns.Count
 If Rs.EOF = False Then

     Sql = "SELECT T_indiceProjet.Client FROM T_indiceProjet "
    Sql = Sql & "WHERE T_indiceProjet.Id=" & Id_IndiceProjet & ";"

    NumErr = 1
    
      Sql = "SELECT T_indiceProjet.Client FROM T_indiceProjet "
    Sql = Sql & "WHERE T_indiceProjet.Id=" & Id_IndiceProjet & ";"

    NumErr = 1

    Set RsComposants = Con.OpenRecordSet(Sql)
    LeCient = UCase(Trim("" & RsComposants!Client))
    Sql = "SELECT  T_Clients.PathComposants FROM T_Clients "
Sql = Sql & "WHERE T_Clients.Client='" & MyReplace(LeCient) & "';"

       Set RsComposants = Con.OpenRecordSet(Sql)
If RsComposants.EOF = False Then
    
    If Trim("" & RsComposants!PathComposants) = "" Then
         PathComposantsDefault = TableauPath.Item("PathComposantsDefault")
   Else
             PathComposantsDefault = RsComposants!PathComposants
'         If Left(PathComposantsDefault, 2) <> "\\" Then PathComposantsDefault = TableauPath.Item("PathServer") & PathComposantsDefault
    
    End If
Else
                 PathComposantsDefault = "" & RsComposants!PathComposants

End If
Else
 PathComposantsDefault = TableauPath.Item("PathComposantsDefault")
End If
PathComposantsDefault = DefinirChemienComplet(TableauPath.Item("PathServer"), PathComposantsDefault)
'If Left(PathComposantsDefault, 2) <> "\\" And Left(PathComposantsDefault, 1) = "\" Then PathComposantsDefault = TableauPath.Item("PathServer") & PathComposantsDefault
'If Right(PathComposantsDefault, 2) = "\\" Then PathComposantsDefault = Mid(PathComposantsDefault, 1, Len(PathComposantsDefault) - 1)

    Dim fs, f, f1, s, sf
 
'  MyExcel.Visible = True
    Set f = Fso.GetFolder(PathComposantsDefault) '\\10.30.0.5\\donnees d entreprise\Utilitaires\cablage\Librairies\COMPOSANTS\")
    Set sf = f.SubFolders
    For Each f1 In sf
       NbColonne = NbColonne + 1
'       MyRange.Application.Visible = True
    MyRange(1, NbColonne) = f1.Name
    MyRange(1, NbColonne).Interior.ColorIndex = 15
    Next
  


If Rs.EOF = True Then Exit Function



TxtPoin = ""
'MyRange(1, NbColonne).AutoFilter
'While Trim(Rep) <> ""
'
'If InStr(1, Trim(Rep), ".") = 0 Then
'    NbColonne = NbColonne + 1
'    MyRange(1, NbColonne) = Rep
'    MyRange(1, NbColonne).Interior.ColorIndex = 15
'
' End If
'    Rep = Dir
'Wend

NbLigne = 1
    
 FormBarGrah.ProgressBar1Caption.Caption = " Exporter liste des Composants :"
 FormBarGrah.ProgressBar1.Value = 0
 If NbLigne = 0 Then NbLigne = 1
 FormBarGrah.ProgressBar1.Max = NbLigne
 IncrmentServer FormBarGrah, ""
Rs.Requery
If Rs.EOF = False Then
'Myrange.Application.Visible = True
    MyRange.Range("A2").CopyFromRecordset Rs
       ReplaceBool MySeet, "A2:A" & MySeet.Range("A1").CurrentRegion.Rows.Count
       
    Row = 2
    FormBarGrah.ProgressBar1.Value = 0
    Set MyRange = MySeet.Range("A1").CurrentRegion
    NbLigne = MyRange.Rows.Count
    If NbLigne = 0 Then NbLigne = 1
    FormBarGrah.ProgressBar1.Max = NbLigne
    FormBarGrah.ProgressBar1Caption.Caption = " Exporter liste des Composants :"
    Set MyRange = MySeet.Range("A1").CurrentRegion
    For Row = 2 To MyRange.Rows.Count
        IncremanteBarGrah FormBarGrah
        IncrmentServer FormBarGrah, ""

    For C = 11 To MyRange.Columns.Count
        If MyRange(1, C) = MyRange(Row, C) Then
        MyRange(Row, C) = 1
        Else
            MyRange(Row, C + 1) = MyRange(Row, C)
            MyRange(Row, C) = 0
        End If
    Next
  
    Next
End If
'Myrange(Row, Myrange.Columns.Count + 1).delte
'While Rs.EOF = False
' IncremanteBarGrah FormBarGrah
' IncrmentServer
'DoEvents
' Myrange(Row, 1) = Replace(Replace("" & Rs!Activer, "Faux", 0), "Vrai", 1)
'  Myrange(Row, 2) = "'" & Rs!DESIGNCOMP
'  Myrange(Row, 3) = "'" & Rs!NUMCOMP
'  Myrange(Row, 4) = "'" & Rs!REFCOMP
'  Myrange(Row, 5) = "'" & Rs!Option
'  Myrange(Row, 6) = "'" & Rs!Code_APP_Lier
'  For I = 7 To Myrange.Columns.Count
'            If Myrange(1, I) = "" & Rs!Path Then Myrange(Row, I) = 1 Else Myrange(Row, I) = 0
'        Next I
'
'
'    Row = Row + 1
'    Rs.MoveNext
'Wend
Fin:
Set MyRange = MySeet.Range("A1").CurrentRegion

Sql = "SELECT  T_indiceProjet.Ensemble, T_indiceProjet.Equipement,[Li] & '_' & [LI_Indice] AS Liste, "
Sql = Sql & "[PI] & '_' & [PI_Indice] AS Piece, T_indiceProjet.Client, T_indiceProjet.CleAc "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Id_IndiceProjet & ";"
Dim RsEntetePage As Recordset
Set RsEntetePage = Con.OpenRecordSet(Sql)

MiseEnPage MySeet, MyRange, "Affaire : " & RsEntetePage!CleAc & Chr(10) & _
 _
     "Pièce : " & RsEntetePage!Piece & Chr(10) & "Liste : " & RsEntetePage!Liste, vbCrLf & "Câblage : " & Replace("" & RsEntetePage!Ensemble, vbCrLf, " ") & Chr(10) & "Equipement : " & Replace("" & RsEntetePage!Equipement, vbCrLf, " ") & Chr(10) & _
     "" _
    , "ENCELADE" & Chr(10) & "Client : " & RsEntetePage!Client & Chr(10) & Format(Date, "dd-mmm-yyyy"), "", "&P/&N", "", 80, "C2", True, 2, True

MaJEncadreXls MyRange, xlThin, xlThin, xlHairline, xlHairline

Set RsEntetePage = Con.CloseRecordSet(RsEntetePage)
Set MyRange = Nothing
Set MySeet = Nothing
End Function
Function OpenModelXlt(Fichier As String) As EXCEL.Workbook
Dim Sql As String
Dim RsSql As Stream
Dim Rs As Recordset
RetournIdApp "EXCEL.EXE"
Set MyExcel = New EXCEL.Application
   
If IsServeur = True Then
    Sql = "SELECT T_Job.Job FROM T_Job "
    Sql = Sql & "WHERE T_Job.Job=" & Command & " AND T_Job.IdExcel<>0;"
    Set Rs = Con.OpenRecordSet(Sql)
    If Rs.EOF = False Then
        Sql = "UPDATE T_Job SET T_Job.IdExcel2 = " & RetournIdApp("EXCEL.EXE", True) & " "
        Sql = Sql & "WHERE T_Job.Job=" & Command & ";"
    Else
         Sql = "UPDATE T_Job SET T_Job.IdExcel = " & RetournIdApp("EXCEL.EXE", True) & " "
        Sql = Sql & "WHERE T_Job.Job=" & Command & ";"
    End If
    Con.Execute Sql
    Set Rs = Con.CloseRecordSet(Rs)
End If
MyExcel.DisplayAlerts = False
'MyExcel.Visible = True
Set OpenModelXlt = MyExcel.Workbooks.Open(Fichier)
End Function

Sub FormatExcelPlage(Plage As Range, Couleur As Long, Merge As Boolean, Grille As Boolean, HorizontalAlignment As Long, VerticalAlignment As Long)
Plage.Interior.ColorIndex = Couleur
If Merge = True Then Plage.Merge
    Plage.HorizontalAlignment = HorizontalAlignment 'xlCenter
    Plage.VerticalAlignment = VerticalAlignment 'xlCenter
If Grille = True Then
    Plage.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Plage.Borders(xlEdgeTop).LineStyle = xlContinuous
    Plage.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Plage.Borders(xlEdgeRight).LineStyle = xlContinuous
    Plage.Borders(xlContinuous).LineStyle = xlContinuous
End If


End Sub

Sub EporteSynthese(Mytype As String, Optional Affaire As Long)

On Error Resume Next
Dim Sql As String
Dim Myrep As String
Dim Rs As Recordset
Dim MyExcel As EXCEL.Application
MyExcel.DisplayAlerts = False
Dim MyWorkbook As EXCEL.Workbook
Dim MySheet As Worksheet
Dim Fso As New FileSystemObject
Dim PathAffair As String
Dim Row As Long
Dim NameFichier
Dim Mabarr As Long
Dim Client As String

RetournIdApp "EXCEL.EXE"
Set MyExcel = New EXCEL.Application
AutoApp.Visible = True
If IsServeur = True Then
    Con.Execute "UPDATE T_Job SET T_Job.IdExcel2 = " & RetournIdApp("EXCEL.EXE", True) & _
    " WHERE T_Job.Job=" & Command & ";"

End If
 Set TableauPath = funPath
'MyExcel.Visible = True
Set MyWorkbook = MyExcel.Workbooks.Add
Set MySheet = IsertSheet(MyWorkbook, "Synthèse", False)
Sql = "SELECT T_Status.Id, T_Status.Status FROM T_Status ORDER BY T_Status.Id;"
Set Rs = Con.OpenRecordSet(Sql)
coll = 1


While Rs.EOF = False

    MySheet.Cells(1, coll) = "" & Rs!Status
      MySheet.Cells(1, coll).Interior.ColorIndex = ChoixCouleur(Val(Rs!Id), True)
       coll = coll + 1
    Rs.MoveNext
Wend
MySheet.Cells(1, coll) = "ARCHIVE"
      MySheet.Cells(1, coll).Interior.ColorIndex = ChoixCouleur(4, True)

'aa = ChoixCouleur(1)
Sql = "SELECT Rq_Synthese_Total.* FROM Rq_Synthese_Total;"
Set Rs = Con.OpenRecordSet(Sql)

    Client = "" & Rs!Client
    If Affaire > 0 Then
        Rs.Filter = "Affaire=" & Affaire
    End If

Row = MySheet.Range(MySheet.Range("A1").CurrentRegion.Address).Rows.Count + 1

For I = 0 To Rs.Fields.Count - 2
    MySheet.Cells(Row, I + 1) = Rs.Fields(I).Name
     MySheet.Cells(Row, I + 1).Interior.ColorIndex = ChoixCouleur(0, True)
Next

Row = MySheet.Range(MySheet.Range("A1").CurrentRegion.Address).Rows.Count
'While Rs.EOF = False
'Mabarr = Mabarr + 1
'Rs.MoveNext
'Wend
FormBarGrah.ProgressBar1.Value = 0
FormBarGrah.ProgressBar1Caption.Caption = Replace(Mytype, "Synt", "Fichier de synthèse ")
IncrmentServer FormBarGrah, "Synthèse"
Rs.Requery
'MySheet.Application.Visible = True
MySheet.Range("A3").CopyFromRecordset Rs
MySheet.Range("A1").CurrentRegion.Replace Chr(13), ""
FormBarGrah.ProgressBar1.Max = MySheet.Range("A3").CurrentRegion.Rows.Count
FormBarGrah.ProgressBar1.Value = 2
For Row = 3 To MySheet.Range("A3").CurrentRegion.Rows.Count
    IncremanteBarGrah FormBarGrah
    IncrmentServer FormBarGrah, "Synthèse"
    MySheet.Rows(Row).Interior.ColorIndex = ChoixCouleur(MySheet.Cells(Row, MySheet.Range("A3").CurrentRegion.Columns.Count), True)
Next
MySheet.Columns(MySheet.Range("A3").CurrentRegion.Columns.Count).Delete
'While Rs.EOF = False
'IncremanteBarGrah FormBarGrah
'IncrmentServer "Synthèse"
'DoEvents
'Row = Row + 1
'    For I = 0 To Rs.Fields.Count - 2
'
'        MySheet.Cells(Row, I + 1) = Rs.Fields(I).Value
'         MySheet.Cells(Row, I + 1).Interior.ColorIndex = ChoixCouleur(Val(Rs!Id), True)
'Next
'Rs.MoveNext
'Wend
'
'sql = "SELECT Rq_Synthese_Total_Archive.* FROM Rq_Synthese_Total_Archive;"
'Set Rs = Con.OpenRecordSet(sql)
'
'    Client = "" & Rs!Client
'    If Affaire > 0 Then
'        Rs.Filter = "Affaire=" & Affaire
'    End If
'
'Row = MySheet.Range(MySheet.Range("A1").CurrentRegion.Address).Rows.Count
'While Rs.EOF = False
'Row = Row + 1
'    For I = 0 To Rs.Fields.Count - 1
'       'MySheet.Cells(Row, I + 1).Select
'        MySheet.Cells(Row, I + 1) = Rs.Fields(I).Value
'         MySheet.Cells(Row, I + 1).Interior.ColorIndex = ChoixCouleur(4, True)
'Next
'Rs.MoveNext
'Wend
MySheet.Range("A2").AutoFilter
'MySheet.Range(MySheet.Range("A1").CurrentRegion.Address).RowHeight = 120
'MySheet.Range(MySheet.Range("A1").CurrentRegion.Address).ColumnWidth = 120
'MySheet.Range(MySheet.Range("A1").CurrentRegion.Address).EntireRow.AutoFit
'MySheet.Range(MySheet.Range("A1").CurrentRegion.Address).EntireColumn.AutoFit
'
MyPath = TableauPath.Item("PathArchiveAutocad")

NameFichier = TableauPath.Item("PathArchiveAutocad")
NameFichier = DefinirChemienComplet(TableauPath.Item("PathServer"), "" & NameFichier)
NameFichier = PathArchive(TableauPath.Item("PathArchiveAutocad"), Client, "" & Affaire, "", Mytype, "Synthèse", 0, 0, 0, 0, True)
NameFichier = NameFichier & ".XLS"
'NameFichier = DefinirChemienComplet(TableauPath.Item("PathServer"), "" & NameFichier)
' If Left(NameFichier, 2) <> "\\" And Left(NameFichier, 1) = "\" Then NameFichier = TableauPath.Item("PathServer") & NameFichier & "\"
'            If Right(NameFichier, 2) = "\\" Then NameFichier = Mid(NameFichier, 1, Len(NameFichier) - 1)
'    If Right(NameFichier, 1) <> "\" Then NameFichier = NameFichier & "\"
'If Affaire > 0 Then
'    NameFichier = NameFichier & PathAffair
'Else
'
'   NameFichier = NameFichier
'End If
'NameFichier = NameFichier & "Synthèse"
'If Affaire > 0 Then
'    NameFichier = NameFichier & "_" & CStr(Affaire)
'End If
'NameFichier = NameFichier & ".XLS"
 If Fso.FileExists(NameFichier) = True Then
    Fso.DeleteFile NameFichier
    End If
   If PortraitPaysage = 0 Then PortraitPaysage = 2
    MiseEnPage MySheet, MySheet.Range("A1").CurrentRegion, "", "", "Date: " & Format(Date, "dd-mmm-yyyy"), "", "&P/&N", "", 80, "A3", True, PortraitPaysage, False, True
    
    MaJEncadreXls MySheet.Range("A1").CurrentRegion, xlThin, xlThin, xlHairline, xlHairline
    
'MyWorkbook.Application.Visible = False
MyWorkbook.Application.DisplayAlerts = False
    MyWorkbook.SaveAs NameFichier, ReadOnlyRecommended:=True
    MyWorkbook.Close False
    MyExcel.Quit
Set MySheet = Nothing
Set MyWorkbook = Nothing
Set MyExcel = Nothing
If IsServeur = True Then
    Con.Execute "UPDATE T_Job SET T_Job.IdExcel2 = 0" & _
    " WHERE T_Job.Job=" & Command & ";"

End If
FormBarGrah.ProgressBar1.Value = 0
FormBarGrah.ProgressBar1.Value = 0
FormBarGrah.ProgressBar1Caption.Caption = " Fin du traitement"
IncrmentServer FormBarGrah, ""
On Error GoTo 0
End Sub
Sub MaJEncadreXls(MyRange As Range, LeftWeight As Long, RightWeight As Long, TopWeight As Long, BottomWeight As Long)
On Error Resume Next

'
' Macro3 Macro
' Macro enregistrée le 14/03/2005 par robert.durupt
'

'
    MyRange.Borders(xlDiagonalDown).LineStyle = xlNone
    MyRange.Borders(xlDiagonalUp).LineStyle = xlNone
    
    MyRange.Borders(xlEdgeLeft).LineStyle = xlContinuous
    MyRange.Borders(xlEdgeLeft).Weight = LeftWeight
   
        MyRange.Borders(xlEdgeRight).LineStyle = xlContinuous
        MyRange.Borders(xlEdgeRight).Weight = RightWeight
        
      
        MyRange.Borders(xlEdgeTop).LineStyle = xlContinuous
        MyRange.Borders(xlEdgeTop).Weight = TopWeight
       
    
        MyRange.Borders(xlEdgeBottom).LineStyle = xlContinuous
        MyRange.Borders(xlEdgeBottom).Weight = BottomWeight
      
    MyRange.Borders(xlInsideVertical).Weight = LeftWeight
     MyRange.Borders(xlInsideHorizontal).Weight = BottomWeight
 On Error GoTo 0
End Sub

