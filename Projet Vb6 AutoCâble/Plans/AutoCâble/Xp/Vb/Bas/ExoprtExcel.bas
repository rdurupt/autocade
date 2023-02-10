Attribute VB_Name = "ExoprtExcel"
Public MyExcel As EXCEL.Application
Public MyWorkbook As EXCEL.Workbook

Public Sub subExporteXls(IdIndiceProjet As Long)
    Dim Rs As Recordset
    Dim PathPl As String
    Dim Sql As String
    Set TableauPath = funPath
    Dim PathPlantVierge As String
'    Set TableauPath = funPath
    Set CollectionTor = Nothing
    Set CollectionTor = New Collection
'
     NUMCOM = 0
    NUMNOTA = 0
    NUMNTOR = 0
     NUMNTORBLOC = 0
    PathArchiveAutocad = TableauPath.Item("PathArchiveAutocad")
     If Left(PathArchiveAutocad, 2) <> "\\" And Left(PathArchiveAutocad, 1) <> "\" Then PathArchiveAutocad = TableauPath.Item("PathServer") & PathArchiveAutocad
     If Right(PathArchiveAutocad, 2) = "\\" Then PathArchiveAutocad = Mid(PathArchiveAutocad, 1, Len(PathArchiveAutocad) - 1)
  
    Sql = "SELECT T_indiceProjet.li FROM T_indiceProjet "
    Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"

    
    Set Rs = Con.OpenRecordSet(Sql)
     PathPlantVierge = TableauPath.Item("PathPlantVierge")
     If Left(PathPlantVierge, 2) <> "\\" And Left(PathPlantVierge, 1) = "\" Then PathPlantVierge = TableauPath.Item("PathServer") & PathPlantVierge
     If Right(PathPlantVierge, 2) = "\\" Then PathPlantVierge = Mid(PathPlantVierge, 1, Len(PathPlantVierge) - 1)
NbError = 0
If Rs.EOF = True Then Exit Sub
    If MsgBox("Voulez-vous exécuter la Macro Exporter Excel" & vbCrLf & Rs!Li, vbQuestion + vbYesNo, "Auto-Câble") = vbNo Then Exit Sub


Sql = "SELECT RqCartouche.* "
Sql = Sql & "FROM RqCartouche "
 Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = False Then
        PathPl = PathArchive(PathArchiveAutocad, "" & Rs!Client, "" & Rs!CleAc, "" & Rs!Pieces, "Li", Rs.Fields("Li"), IdIndiceProjet, Rs.Fields("PI_Indice"), Rs.Fields("LI_Indice"), Rs!Version)
   
    ExporteXls PathPl, IdIndiceProjet, PathPl

    End If
        DoEvents

Fin:
    ReDim TableauDeConnecteurs(0)
    
    
     FormBarGrah.ProgressBar1.Value = 0
     FormBarGrah.ProgressBar1Caption.Caption = " Fin du traitement"
    
MsgBox "Fin du traitement" & vbCrLf & NbError & " Erreur(s) détectée(s) !"
MenuShow = False
End Sub

Public Sub ExporteXls(Xls As String, IdIndiceProjet As Long, Optional PathPl As String)
Dim Fso As New FileSystemObject
Dim Sql As String
Dim RsIdProjet As Recordset
Dim Rs As Recordset
Dim PathModelXls As String
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
         If Left(DbCatalogue, 2) <> "\\" And Left(DbCatalogue, 1) = "\" Then DbCatalogue = TableauPath.Item("PathServer") & DbCatalogue
            If Right(DbCatalogue, 2) = "\\" Then DbCatalogue = Mid(DbCatalogue, 1, Len(DbCatalogue) - 1)
    
    End If
Else
    DbCatalogue = ""
End If





PathModelXls = TableauPath.Item("PathModelXls")
         If Left(PathModelXls, 2) <> "\\" And Left(PathModelXls, 1) = "\" Then PathModelXls = TableauPath.Item("PathServer") & PathModelXls
          If Right(PathModelXls, 2) = "\\" Then PathModelXls = Mid(PathModelXls, 1, Len(PathModelXls) - 1)

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
'***********************************************************************************************************************
'*                                      Exporte la liste des T_Noeuds.                                              *

    Sql = "SELECT T_Noeuds.ACTIVER, T_Noeuds.Fleche_Droite, T_Noeuds.TORON_PRINCIPAL, T_Noeuds.NŒUDS, T_Noeuds.LONGUEUR,  "
    Sql = Sql & "T_Noeuds.LONGUEUR_CUMULEE, T_Noeuds.DESIGN_HAB, T_Noeuds.CODE_RSA, T_Noeuds.CODE_PSA,  "
    Sql = Sql & "T_Noeuds.CODE_ENC, T_Noeuds.DIAMETRE, T_Noeuds.CLASSE_T  "
    Sql = Sql & "FROM T_Noeuds "
    Sql = Sql & "WHERE T_Noeuds.Id_IndiceProjet = " & IdIndiceProjet & " "
    Sql = Sql & "ORDER BY T_Noeuds.NŒUDS;"
    Set Rs = Con.OpenRecordSet(Sql)

    ExporteXlsNoeuds Rs, IdIndiceProjet
    
'***********************************************************************************************************************
'*                                      Exporte la liste des Critères.                                              *

    Sql = "SELECT T_Critères.ACTIVER, T_Critères.CODE_CRITERE, T_Critères.CRITERES FROM T_Critères "
    Sql = Sql & "WHERE T_Critères.Id_IndiceProjet = " & IdIndiceProjet & " "
    Sql = Sql & "ORDER BY T_Critères.CODE_CRITERE;"
    Set Rs = Con.OpenRecordSet(Sql)

    ExporteXlsCriteres Rs, IdIndiceProjet
'***********************************************************************************************************************
'*                                      Exporte la liste des connecteurs.                                              *

    Sql = "SELECT Connecteurs.ACTIVER, Connecteurs.CONNECTEUR, Connecteurs.[O/N], Connecteurs.DESIGNATION, Connecteurs.CODE_APP,  "
    Sql = Sql & "Connecteurs.N°, Connecteurs.POS, Connecteurs.[POS-OUT], Connecteurs.PRECO1, Connecteurs.PRECO2,  "
    Sql = Sql & "Connecteurs.OPTION, Connecteurs.[100%], Connecteurs.Pylone, Connecteurs.Colonne, Connecteurs.Ligne  "
    Sql = Sql & "FROM Connecteurs "
    Sql = Sql & "WHERE Connecteurs.Id_IndiceProjet = " & IdIndiceProjet & " "
    Sql = Sql & "ORDER BY Connecteurs.N°;"
    Set Rs = Con.OpenRecordSet(Sql)

    ExporteXlsConnecteur Rs, IdIndiceProjet
'***********************************************************************************************************************
'*                                      Exporte la liste des Fils.                                                     *
    
    Sql = "SELECT Ligne_Tableau_fils.ACTIVER, Ligne_Tableau_fils.LIAI, Ligne_Tableau_fils.DESIGNATION, Ligne_Tableau_fils.FIL,  "
   Sql = Sql & "Ligne_Tableau_fils.SECT, Ligne_Tableau_fils.TEINT, Ligne_Tableau_fils.TEINT2, Ligne_Tableau_fils.ISO,  "
   Sql = Sql & "Ligne_Tableau_fils.LONG, Ligne_Tableau_fils.[LONG CP], Ligne_Tableau_fils.COUPE, Ligne_Tableau_fils.POS,  "
   Sql = Sql & "Ligne_Tableau_fils.[POS-OUT], Ligne_Tableau_fils.FA, Ligne_Tableau_fils.APP, Ligne_Tableau_fils.VOI, Ligne_Tableau_fils.POS2,  "
   Sql = Sql & "Ligne_Tableau_fils.[POS-OUT2], Ligne_Tableau_fils.FA2, Ligne_Tableau_fils.APP2, Ligne_Tableau_fils.VOI2,  "
   Sql = Sql & "Ligne_Tableau_fils.PRECO, Ligne_Tableau_fils.OPTION "
   Sql = Sql & "FROM Ligne_Tableau_fils WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & IdIndiceProjet & " ORDER BY val('' & Ligne_Tableau_fils.FIL);"
    Set Rs = Con.OpenRecordSet(Sql)
    ExporteXlsFils Rs, IdIndiceProjet
    
'***********************************************************************************************************************
'*                                      Exporte la liste des Composants.                                              *
       
    Sql = "SELECT Composants.*  "
    Sql = Sql & "FROM Composants "
    Sql = Sql & "WHERE Composants.Id_IndiceProjet = " & IdIndiceProjet & " "
    Sql = Sql & "ORDER BY Composants.NUMCOMP;"
    Set Rs = Con.OpenRecordSet(Sql)

    ExporteXlsComposants Rs, IdIndiceProjet
    
'***********************************************************************************************************************
'*                                      Exporte la liste des Notas.                                                    *
    
    Sql = "SELECT Nota.ACTIVER, Nota.NOTA, Nota.NUMNOTA FROM Nota "
    Sql = Sql & "WHERE Nota.Id_IndiceProjet= " & IdIndiceProjet & " "
    Sql = Sql & "ORDER BY Nota.NUMNOTA ;"

    Set Rs = Con.OpenRecordSet(Sql)
    ExporteXlsNotas Rs, IdIndiceProjet
    
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
    NomenclatureOk = Nomenclature(IdIndiceProjet, PathPl)

End If
'***********************************************************************************************************************
'*                                      Exporte RAPPORT DE_CONTRÔLE_FILAIRE.                                            *
ExporteXlsRAPPORT_DE_CONTROLE_FILAIRE IdIndiceProjet
'***********************************************************************************************************************
'*                                      Exporte Fiche de Contrôle.                                            *
ExporteXlsFiche_de_Controle IdIndiceProjet
'***********************************************************************************************************************
'*                                      Supprime le fichier Excel s'il existe                                          *
If Fso.FileExists(Xls & ".xls") Then Fso.DeleteFile Xls & ".xls"
Set Fso = Nothing
On Error Resume Next
'***********************************************************************************************************************
'*                                      Enregistre le fichier & referme Excel.                                         *
MyWorkbook.Worksheets(1).Select

MyWorkbook.SaveAs Xls
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
If Err Then MsgBox Err.Description
On Error GoTo 0
MyWorkbook.Close False
Set MyWorkbook = Nothing
MyExcel.Quit

Set MyExcel = Nothing
'***********************************************************************************************************************
 FormBarGrah.ProgressBar1.Value = 0
 FormBarGrah.ProgressBar1Caption = " Fin du traitement:"

End Sub







Function ExporteXlsPrixFils(Rs As Recordset, Id_IndiceProjet As Long)
Dim MySeet As EXCEL.Worksheet
Dim Myrange As EXCEL.Range
Dim Row As Long
Dim NbLigne As Long
DoEvents
    NbLigne = 0
If Rs.EOF = True Then Exit Function
While Rs.EOF = False
NbLigne = NbLigne + 1
Rs.MoveNext
Wend

Rs.MoveFirst
'MyWorkbook.Application.Visible = True
Set MySeet = IsertSheet(MyWorkbook, "Prix Fils", True)

DeleteRow MySeet, True

Set Myrange = MySeet.Range("A5").CurrentRegion
For i = 0 To Rs.Fields.Count - 2
    Myrange(1, i + 1) = Rs.Fields(i).Name
Next
Set Myrange = MySeet.Range("A5").CurrentRegion

    Myrange.Interior.ColorIndex = 15
    Myrange.HorizontalAlignment = xlCenter
        
Row = 2
 FormBarGrah.ProgressBar1.Value = 0
 FormBarGrah.ProgressBar1.Max = NbLigne
 FormBarGrah.ProgressBar1Caption.Caption = " Exporter Prix du Câble :"
While Rs.EOF = False
     IncremanteBarGrah FormBarGrah
    DoEvents
    Myrange(Row, 1) = "" & Rs!TEINT
    Myrange(Row, 2) = "" & Rs!ISO
    Myrange(Row, 3) = Val(Replace("" & Rs!SECT, ",", "."))
    Myrange(Row, 4) = Val(Replace("" & Rs!Longeur, ",", "."))
    Myrange(Row, 5) = Val(Replace("" & Rs![Prix u], ",", "."))
    Myrange(Row, 6).FormulaR1C1 = "" & Rs![Prix Total]
    
    Rs.MoveNext
    Row = Row + 1
Wend
Dim Sql As String
Set Myrange = MySeet.Range("A5").CurrentRegion

Sql = "SELECT  T_indiceProjet.Ensemble, T_indiceProjet.Equipement,[Li] & '_' & [LI_Indice] AS Liste, "
Sql = Sql & "[PI] & '_' & [PI_Indice] AS Piece, T_indiceProjet.Client, T_indiceProjet.CleAc "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Id_IndiceProjet & ";"
Dim RsEntetePage As Recordset
Set RsEntetePage = Con.OpenRecordSet(Sql)

MySeet.Range("E2") = "SOUS TOTAL"
FormatExcelPlage MySeet.Range("E2"), 15, False, True, xlCenter, xlCenter
r1 = MySeet.Range(Myrange(2, Myrange.Columns.Count).Address).Row
r2 = MySeet.Range(Myrange(Myrange.Rows.Count, Myrange.Columns.Count).Address).Row
MySeet.Range("F2").FormulaR1C1 = "=SUBTOTAL(9,R[" & r1 - 2 & "]C:R[" & r2 - 2 & "]C)"
FormatExcelPlage MySeet.Range("F2"), 2, False, True, xlCenter, xlCenter

MiseEnPage MySeet, Myrange, "Affaire : " & RsEntetePage!CleAc & vbCrLf & _
 _
     "Pièce : " & RsEntetePage!Piece & vbCrLf & "Liste : " & RsEntetePage!Liste, vbCrLf & "Câblage : " & Replace("" & RsEntetePage!Ensemble, vbCrLf, " ") & vbCrLf & "Equipement : " & Replace("" & RsEntetePage!Equipement, vbCrLf, " ") & vbCrLf & _
     "" _
    , "ENCELADE" & vbCrLf & "Client : " & RsEntetePage!Client & vbCrLf & Format(Date, "dd-mmm-yyyy"), "", "&P/&N", "", 80, "B6", True, 1, True
    
      MaJEncadreXls Myrange, xlThin, xlThin, xlHairline, xlHairline
      
Set RsEntetePage = Con.CloseRecordSet(RsEntetePage)
Set Myrange = Nothing
Set MySeet = Nothing

End Function
Function ExporteXlsHabillages(Rs As Recordset, Id_IndiceProjet As Long)
Dim MySeet As EXCEL.Worksheet
Dim Myrange As EXCEL.Range
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

Set Myrange = MySeet.Range("A1").CurrentRegion
'Myrange.Application.Visible = True
Row = 6
 FormBarGrah.ProgressBar1.Value = 0
 FormBarGrah.ProgressBar1.Max = NbLigne
 FormBarGrah.ProgressBar1Caption.Caption = " Exporter liste des Habilage :"
 For i = 0 To Rs.Fields.Count - 2
    MySeet.Cells(5, i + 1) = Rs.Fields(i).Name

 Next
While Rs.EOF = False
     IncremanteBarGrah FormBarGrah
     For i = 0 To Rs.Fields.Count - 2
  
     
         If Rs(i).Name = "Prix Total" Then
            MySeet.Cells(Row, i + 1).FormulaR1C1 = "=(RC[-1]*RC[-2])"
         End If
 Next
    DoEvents
    For i = 0 To Rs.Fields.Count - 2
      If Rs(i).Name = "Prix Total" Then
                                
    Else
        MySeet.Cells(Row, i + 1) = Trim(Replace("" & Rs(i), vbCrLf, ""))
    End If
    Next
    Row = Row + 1
    Rs.MoveNext
Wend
Dim Sql As String
Set Myrange = MySeet.Range("A5").CurrentRegion

Sql = "SELECT  T_indiceProjet.Ensemble, T_indiceProjet.Equipement,[Li] & '_' & [LI_Indice] AS Liste, "
Sql = Sql & "[PI] & '_' & [PI_Indice] AS Piece, T_indiceProjet.Client, T_indiceProjet.CleAc "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Id_IndiceProjet & ";"
Dim RsEntetePage As Recordset
Set RsEntetePage = Con.OpenRecordSet(Sql)

MySeet.Range("C2") = "SOUS TOTAL"
FormatExcelPlage MySeet.Range("C2"), 15, False, True, xlCenter, xlCenter
r1 = MySeet.Range(Myrange(2, Myrange.Columns.Count).Address).Row
r2 = MySeet.Range(Myrange(Myrange.Rows.Count, Myrange.Columns.Count).Address).Row
MySeet.Range("D2").FormulaR1C1 = "=SUBTOTAL(9,R[" & r1 - 2 & "]C:R[" & r2 - 2 & "]C)"
FormatExcelPlage MySeet.Range("D2"), 2, False, True, xlCenter, xlCenter

MiseEnPage MySeet, Myrange, "Affaire : " & RsEntetePage!CleAc & vbCrLf & _
 _
     "Pièce : " & RsEntetePage!Piece & vbCrLf & "Liste : " & RsEntetePage!Liste, vbCrLf & "Câblage : " & Replace("" & RsEntetePage!Ensemble, vbCrLf, " ") & vbCrLf & "Equipement : " & Replace("" & RsEntetePage!Equipement, vbCrLf, " ") & vbCrLf & _
     "" _
    , "ENCELADE" & vbCrLf & "Client : " & RsEntetePage!Client & vbCrLf & Format(Date, "dd-mmm-yyyy"), "", "&P/&N", "", 80, "B6", True, 1, True
    
      MaJEncadreXls Myrange, xlThin, xlThin, xlHairline, xlHairline
      
Set RsEntetePage = Con.CloseRecordSet(RsEntetePage)
Set Myrange = Nothing
Set MySeet = Nothing

    
End Function


Function ExporteXlsFils(Rs As Recordset, Id_IndiceProjet As Long)
Dim MySeet As EXCEL.Worksheet
Dim Myrange As EXCEL.Range
Dim Row As Long
Dim NbLigne As Long
DoEvents
    NbLigne = 0
If Rs.EOF = True Then Exit Function
While Rs.EOF = False
NbLigne = NbLigne + 1
Rs.MoveNext
Wend
Rs.MoveFirst

Set MySeet = IsertSheet(MyWorkbook, "Ligne_Tableau_fils", True)
IsertColonne "ACTIVER", MySeet
DeleteRow MySeet, True

Set Myrange = MySeet.Range("A1").CurrentRegion
ExcelCreatTitre Myrange, Rs
Row = 2
 FormBarGrah.ProgressBar1.Value = 0
 FormBarGrah.ProgressBar1.Max = NbLigne
 FormBarGrah.ProgressBar1Caption.Caption = " Exporter liste des Fils :"
While Rs.EOF = False
     IncremanteBarGrah FormBarGrah
    DoEvents
    ExcelCreatTitre Myrange(Row, 1), Rs, True
'     Myrange(Row, 1) = Replace(Replace("" & Rs!ACTIVER, "Faux", 0), "Vrai", 1)
'     Myrange(Row, 2) = "'" & Rs!Liai
'    Myrange(Row, 3) = "'" & Rs!DESIGNATION
'    Myrange(Row, 4) = "'" & Rs!Fil
'    Myrange(Row, 5) = "'" & Rs!SECT
'    Myrange(Row, 6) = "'" & Rs!TEINT
'    Myrange(Row, 7) = "'" & Rs!TEINT2
'    Myrange(Row, 8) = "'" & Rs!ISO
'    Myrange(Row, 9) = "'" & Rs!Long
'    Myrange(Row, 10) = "'" & Rs![LONG CP]
'    Myrange(Row, 11) = "'" & Rs!Coupe
'    Myrange(Row, 12) = "'" & Rs!POS
'    Myrange(Row, 13) = "'" & Rs![POS-OUT]
'    Myrange(Row, 14) = "'" & Rs!FA
'    Myrange(Row, 15) = "'" & Rs![App]
'    Myrange(Row, 16) = "'" & Rs!VOI
'
'    Myrange(Row, 17) = "'" & Rs![POS2]
'    Myrange(Row, 18) = "'" & Rs![POS-OUT2]
'    Myrange(Row, 19) = "'" & Rs![FA2]
'
'    Myrange(Row, 20) = "'" & Rs![app2]
'    Myrange(Row, 21) = "'" & Rs![VOI2]
'    Myrange(Row, 22) = "'" & Rs![PRECO]
'    Myrange(Row, 23) = "'" & Rs![Option]
    Rs.MoveNext
    Row = Row + 1
Wend
Dim Sql As String
Set Myrange = MySeet.Range("A1").CurrentRegion

Sql = "SELECT  T_indiceProjet.Ensemble, T_indiceProjet.Equipement,[Li] & '_' & [LI_Indice] AS Liste, "
Sql = Sql & "[PI] & '_' & [PI_Indice] AS Piece, T_indiceProjet.Client, T_indiceProjet.CleAc "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Id_IndiceProjet & ";"
Dim RsEntetePage As Recordset
Set RsEntetePage = Con.OpenRecordSet(Sql)

MiseEnPage MySeet, Myrange, "Affaire : " & RsEntetePage!CleAc & vbCrLf & _
 _
     "Pièce : " & RsEntetePage!Piece & vbCrLf & "Liste : " & RsEntetePage!Liste, vbCrLf & "Câblage : " & Replace("" & RsEntetePage!Ensemble, vbCrLf, " ") & vbCrLf & "Equipement : " & Replace("" & RsEntetePage!Equipement, vbCrLf, " ") & vbCrLf & _
      "" _
    , "ENCELADE" & vbCrLf & "Client : " & RsEntetePage!Client & vbCrLf & Format(Date, "dd-mmm-yyyy"), "", "&P/&N", "", 80, "c2", True, 2, True

  MaJEncadreXls Myrange, xlThin, xlThin, xlHairline, xlHairline
  
Set RsEntetePage = Con.CloseRecordSet(RsEntetePage)
Set Myrange = Nothing
Set MySeet = Nothing
End Function
Function ExporteXlsNoeuds(Rs As Recordset, Id_IndiceProjet As Long)

Dim MySeet As EXCEL.Worksheet
Dim Myrange As EXCEL.Range
Dim Row As Long
Dim NbLigne As Long
    NbLigne = 0
    If Rs.EOF = True Then Exit Function
While Rs.EOF = False
NbLigne = NbLigne + 1
Rs.MoveNext
Wend
Rs.Requery
'MyWorkbook.Application.Visible = True
Set MySeet = IsertSheet(MyWorkbook, "NOEUDS", True)
IsertColonne "ACTIVER", MySeet
DeleteRow MySeet, True
Set Myrange = MySeet.Range("A1").CurrentRegion
ExcelCreatTitre Myrange, Rs
Row = 2
 FormBarGrah.ProgressBar1.Value = 0
 FormBarGrah.ProgressBar1.Max = NbLigne
 FormBarGrah.ProgressBar1Caption.Caption = " Exporter liste des NOEUDS :"
While Rs.EOF = False
 IncremanteBarGrah FormBarGrah
DoEvents
ExcelCreatTitre Myrange(Row, 1), Rs, True
'    Myrange(Row, 1) = Replace("" & Rs!Fleche_Droite, "Vrai", 1)
'  If Myrange(Row, 1) <> 1 Then Myrange(Row, 1) = 0
'
' Myrange(Row, 2) = Replace("" & Rs!TORON_PRINCIPAL, "Vrai", 1)
'  If Myrange(Row, 2) <> 1 Then Myrange(Row, 2) = 0
'    Myrange(Row, 3) = Replace("" & Rs!ACTIVER, "Vrai", 1)
'  If Myrange(Row, 3) <> 1 Then Myrange(Row, 3) = 0
'  Myrange(Row, 4) = "'" & Rs!Noeuds
'  Myrange(Row, 5) = Val("" & Rs!Longueur)
'  Myrange(Row, 6) = Val("" & Rs!LONGUEUR_CUMULEE)
'  Myrange(Row, 7) = "" & Rs!DESIGN_HAB
'  Myrange(Row, 8) = "" & Rs!CODE_RSA
'  Myrange(Row, 9) = "" & Rs!CODE_PSA
'  Myrange(Row, 10) = "" & Rs!CODE_ENC
'  Myrange(Row, 11) = "" & Rs!DIAMETRE
'  Myrange(Row, 12) = "" & Rs!CLASSE_T
    Rs.MoveNext
    Row = Row + 1
Wend

Set Myrange = MySeet.Range("A1").CurrentRegion
Dim Sql As String
Sql = "SELECT  T_indiceProjet.Ensemble, T_indiceProjet.Equipement,[Li] & '_' & [LI_Indice] AS Liste, "
Sql = Sql & "[PI] & '_' & [PI_Indice] AS Piece, T_indiceProjet.Client, T_indiceProjet.CleAc "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Id_IndiceProjet & ";"
Dim RsEntetePage As Recordset

Set RsEntetePage = Con.OpenRecordSet(Sql)

MiseEnPage MySeet, Myrange, "Affaire : " & RsEntetePage!CleAc & vbCrLf & _
 _
     "Pièce : " & RsEntetePage!Piece & vbCrLf & "Liste : " & RsEntetePage!Liste, vbCrLf & "Câblage : " & Replace("" & RsEntetePage!Ensemble, vbCrLf, " ") & vbCrLf & "Equipement : " & Replace("" & RsEntetePage!Equipement, vbCrLf, " ") & vbCrLf & _
     "" _
    , "ENCELADE" & vbCrLf & "Client : " & RsEntetePage!Client & vbCrLf & Format(Date, "dd-mmm-yyyy"), "", "&P/&N", "", 80, "c2", True, 2, True
    
      MaJEncadreXls Myrange, xlThin, xlThin, xlHairline, xlHairline

Set RsEntetePage = Con.CloseRecordSet(RsEntetePage)
Set Myrange = Nothing
Set MySeet = Nothing
End Function

Function ExporteXlsCriteres(Rs As Recordset, Id_IndiceProjet As Long)
Dim MySeet As EXCEL.Worksheet
Dim Myrange As EXCEL.Range
Dim Row As Long
Dim NbLigne As Long
    NbLigne = 0
    If Rs.EOF = True Then Exit Function
While Rs.EOF = False
NbLigne = NbLigne + 1
Rs.MoveNext
Wend
Rs.Requery
'MyWorkbook.Application.Visible = True
Set MySeet = IsertSheet(MyWorkbook, "Critères", True)
IsertColonne "ACTIVER", MySeet
DeleteRow MySeet, True

Set Myrange = MySeet.Range("A1").CurrentRegion
ExcelCreatTitre Myrange, Rs
Row = 2
 FormBarGrah.ProgressBar1.Value = 0
 FormBarGrah.ProgressBar1.Max = NbLigne
 FormBarGrah.ProgressBar1Caption.Caption = " Exporter liste des Critères :"
While Rs.EOF = False
 IncremanteBarGrah FormBarGrah
DoEvents
ExcelCreatTitre Myrange(Row, 1), Rs, True

'     Myrange(Row, 1) = Replace(Replace("" & Rs!ACTIVER, "Faux", 0), "Vrai", 1)
'  Myrange(Row, 2) = "'" & Rs!CODE_CRITERE
'  Myrange(Row, 3) = "'" & Rs!CRITERES
  
    Rs.MoveNext
    Row = Row + 1
Wend

Set Myrange = MySeet.Range("A1").CurrentRegion
Dim Sql As String
Sql = "SELECT  T_indiceProjet.Ensemble, T_indiceProjet.Equipement,[Li] & '_' & [LI_Indice] AS Liste, "
Sql = Sql & "[PI] & '_' & [PI_Indice] AS Piece, T_indiceProjet.Client, T_indiceProjet.CleAc "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Id_IndiceProjet & ";"
Dim RsEntetePage As Recordset

Set RsEntetePage = Con.OpenRecordSet(Sql)

MiseEnPage MySeet, Myrange, "Affaire : " & RsEntetePage!CleAc & vbCrLf & _
 _
     "Pièce : " & RsEntetePage!Piece & vbCrLf & "Liste : " & RsEntetePage!Liste, vbCrLf & "Câblage : " & Replace("" & RsEntetePage!Ensemble, vbCrLf, " ") & vbCrLf & "Equipement : " & Replace("" & RsEntetePage!Equipement, vbCrLf, " ") & vbCrLf & _
     "" _
    , "ENCELADE" & vbCrLf & "Client : " & RsEntetePage!Client & vbCrLf & Format(Date, "dd-mmm-yyyy"), "", "&P/&N", "", 80, "C2", True, 2, True

MaJEncadreXls Myrange, xlThin, xlThin, xlHairline, xlHairline

Set RsEntetePage = Con.CloseRecordSet(RsEntetePage)
Set Myrange = Nothing
Set MySeet = Nothing

End Function
Public Sub IsertColonne(NameCol As String, Sheet As Worksheet)
Dim Myrange As Range
Dim Trouve As Boolean
Trouve = False
Set Myrange = Sheet.Range("A1").CurrentRegion
For i = 1 To Myrange.Columns.Count
    If UCase(Myrange(1, i)) = UCase(NameCol) Then
        Trouve = True
        Exit For
    End If
Next i
If Trouve = False Then
    Sheet.Range("A1").Insert Shift:=xlToRight
    Sheet.Range("A1") = NameCol
    Sheet.Range("A1").Interior.ColorIndex = 15
     Sheet.Range("A1").HorizontalAlignment = xlCenter
    
    Sheet.Range("A1").AutoFilter
   Sheet.Range("A1").AutoFilter
End If
    
End Sub
Function ExporteXlsRAPPORT_DE_CONTROLE_FILAIRE(Id_IndiceProjet As Long)
Dim MySeet As EXCEL.Worksheet
Dim Myrange As EXCEL.Range
Dim Row As Long
Set MySeet = IsertSheet(MyWorkbook, "RAPPORT DE CONTRÔLE CONTINUITE", True)
Set Myrange = MySeet.Range("A1").CurrentRegion

'Myrange.Application.Visible = True
Dim Sql As String
Sql = "SELECT  T_indiceProjet.Ensemble, T_indiceProjet.Equipement,[Li] & '_' & [LI_Indice] AS Liste, "
Sql = Sql & "[PI] & '_' & [PI_Indice] AS Piece, T_indiceProjet.Client, T_indiceProjet.CleAc "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Id_IndiceProjet & ";"
Dim RsEntetePage As Recordset

Set RsEntetePage = Con.OpenRecordSet(Sql)

MiseEnPage MySeet, Myrange, "Affaire : " & RsEntetePage!CleAc & vbCrLf & _
 _
     "Pièce : " & RsEntetePage!Piece & vbCrLf & "Liste : " & RsEntetePage!Liste, vbCrLf & "Câblage : " & Replace("" & RsEntetePage!Ensemble, vbCrLf, " ") & vbCrLf & "Equipement : " & Replace("" & RsEntetePage!Equipement, vbCrLf, " ") & vbCrLf & _
     "" _
    , "ENCELADE" & vbCrLf & "Client : " & RsEntetePage!Client & vbCrLf & Format(Date, "dd-mmm-yyyy"), "", "&P/&N", "", 80, "A2", True, 2, False, False, False, 2.5, False, False
    
MaJEncadreXls Myrange, xlThin, xlThin, xlHairline, xlHairline

Set RsEntetePage = Con.CloseRecordSet(RsEntetePage)
Set Myrange = Nothing
Set MySeet = Nothing
End Function
Function ExporteXlsFiche_de_Controle(Id_IndiceProjet As Long)
Dim MySeet As EXCEL.Worksheet
Dim Myrange As EXCEL.Range
Dim Row As Long
Set MySeet = IsertSheet(MyWorkbook, "Fiche de Contrôle", True)
Set Myrange = MySeet.Range("A1").CurrentRegion


Set Myrange = MySeet.Range("A1").CurrentRegion
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

MiseEnPage MySeet, Myrange, "Affaire : " & RsEntetePage!CleAc & vbCrLf & _
 _
     "Pièce : " & RsEntetePage!Piece & vbCrLf & "Liste : " & RsEntetePage!Liste, vbCrLf & "Câblage : " & Replace("" & RsEntetePage!Ensemble, vbCrLf, " ") & vbCrLf & "Equipement : " & Replace("" & RsEntetePage!Equipement, vbCrLf, " ") & vbCrLf & _
     "" _
    , "ENCELADE" & vbCrLf & "Client : " & RsEntetePage!Client & vbCrLf & Format(Date, "dd-mmm-yyyy"), "", "" & MyPiedTxt, "&P/&N", 100, "A2", True, xlPortrait, False, False, False, 2.5, False, False
    
'MaJEncadreXls Myrange, xlThin, xlThin, xlHairline, xlHairline

Set RsEntetePage = Con.CloseRecordSet(RsEntetePage)
Set Myrange = Nothing
Set MySeet = Nothing
End Function


Function ExporteXlsConnecteur(Rs As Recordset, Id_IndiceProjet As Long)
Dim MySeet As EXCEL.Worksheet
Dim Myrange As EXCEL.Range
Dim Row As Long
Dim NbLigne As Long
    NbLigne = 0
    If Rs.EOF = True Then Exit Function
While Rs.EOF = False
NbLigne = NbLigne + 1
Rs.MoveNext
Wend
Rs.Requery
Set MySeet = IsertSheet(MyWorkbook, "Connecteurs", True)
IsertColonne "ACTIVER", MySeet
DeleteRow MySeet, True

Set Myrange = MySeet.Range("A1").CurrentRegion
ExcelCreatTitre Myrange, Rs
Row = 2
 FormBarGrah.ProgressBar1.Value = 0
 FormBarGrah.ProgressBar1.Max = NbLigne
 FormBarGrah.ProgressBar1Caption.Caption = " Exporter liste des Connecteurs :"
While Rs.EOF = False
 IncremanteBarGrah FormBarGrah
DoEvents
'Myrange.Application.Visible = True
ExcelCreatTitre Myrange(Row, 1), Rs, True
' Myrange(Row, 1) = Replace(Replace("" & Rs!ACTIVER, "Faux", 0), "Vrai", 1)
'  Myrange(Row, 2) = "'" & Rs!Connecteur
'  Myrange(Row, 3) = Abs(Rs![O/N])
'  Myrange(Row, 4) = "'" & Rs!DESIGNATION
'  Myrange(Row, 5) = "'" & Rs!Code_APP
'  Myrange(Row, 6) = "'" & Rs![N°]
'  Myrange(Row, 7) = "'" & Rs!POS
'  Myrange(Row, 8) = "'" & Rs![POS-OUT]
'  Myrange(Row, 9) = "'" & Rs!PRECO1
'  Myrange(Row, 10) = "'" & Rs!PRECO2
'  Myrange(Row, 11) = "'" & Rs![Option]
'  Myrange(Row, 12) = "'" & Rs![100%]
'  Myrange(Row, 13) = "'" & Rs![Pylone]
'  Myrange(Row, 14) = "'" & Rs![Colonne]
' Myrange(Row, 15) = "'" & Rs![Ligne]
    Rs.MoveNext
    Row = Row + 1
Wend

Set Myrange = MySeet.Range("A1").CurrentRegion
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


MiseEnPage MySeet, Myrange, "Affaire : " & RsEntetePage!CleAc & vbCrLf & _
 _
     "Pièce : " & RsEntetePage!Piece & vbCrLf & "Liste : " & RsEntetePage!Liste, vbCrLf & "Câblage : " & Replace("" & RsEntetePage!Ensemble, vbCrLf, " ") & vbCrLf & "Equipement : " & Replace("" & RsEntetePage!Equipement, vbCrLf, " ") & vbCrLf & _
     "" _
    , "ENCELADE" & vbCrLf & "Client : " & RsEntetePage!Client & vbCrLf & Format(Date, "dd-mmm-yyyy"), "", "" & TxtPieCentre, "&P/&N", 80, "C2", True, 2, True, , , 2.5

MaJEncadreXls Myrange, xlThin, xlThin, xlHairline, xlHairline

Set RsEntetePage = Con.CloseRecordSet(RsEntetePage)
Set Myrange = Nothing
Set MySeet = Nothing
End Function
Function ExporteXlsNotas(Rs As Recordset, Id_IndiceProjet As Long)
Dim MySeet As EXCEL.Worksheet
Dim Myrange As EXCEL.Range
Dim Row As Long
Dim NbLigne As Long
    NbLigne = 0
    If Rs.EOF = True Then Exit Function
While Rs.EOF = False
NbLigne = NbLigne + 1
Rs.MoveNext
Wend
Rs.MoveFirst
Set MySeet = IsertSheet(MyWorkbook, "Notas", True)
IsertColonne "ACTIVER", MySeet
DeleteRow MySeet, True

Set Myrange = MySeet.Range("A1").CurrentRegion
ExcelCreatTitre Myrange, Rs
Row = 2
 FormBarGrah.ProgressBar1.Value = 0
 FormBarGrah.ProgressBar1.Max = NbLigne
 FormBarGrah.ProgressBar1Caption.Caption = " Exporter liste des Connecteurs :"
While Rs.EOF = False
 IncremanteBarGrah FormBarGrah
DoEvents
ExcelCreatTitre Myrange(Row, 1), Rs, True
'     Myrange(Row, 1) = Replace(Replace("" & Rs!ACTIVER, "Faux", 0), "Vrai", 1)
'  Myrange(Row, 2) = "'" & Rs!Nota
'  Myrange(Row, 3) = "'" & Rs!NUMNOTA
 
    Rs.MoveNext
    Row = Row + 1
Wend
Set Myrange = MySeet.Range("A1").CurrentRegion
Dim Sql As String
Sql = "SELECT  T_indiceProjet.Ensemble, T_indiceProjet.Equipement,[Li] & '_' & [LI_Indice] AS Liste, "
Sql = Sql & "[PI] & '_' & [PI_Indice] AS Piece, T_indiceProjet.Client, T_indiceProjet.CleAc "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Id_IndiceProjet & ";"
Dim RsEntetePage As Recordset
Set RsEntetePage = Con.OpenRecordSet(Sql)

MiseEnPage MySeet, Myrange, "Affaire : " & RsEntetePage!CleAc & vbCrLf & _
 _
     "Pièce : " & RsEntetePage!Piece & vbCrLf & "Liste : " & RsEntetePage!Liste, vbCrLf & "Câblage : " & Replace("" & RsEntetePage!Ensemble, vbCrLf, " ") & vbCrLf & "Equipement : " & Replace("" & RsEntetePage!Equipement, vbCrLf, " ") & vbCrLf & _
     "" _
    , "ENCELADE" & vbCrLf & "Client : " & RsEntetePage!Client & vbCrLf & Format(Date, "dd-mmm-yyyy"), "", "&P/&N", "", 80, "C2", True, 2

MaJEncadreXls Myrange, xlThin, xlThin, xlHairline, xlHairline

Set RsEntetePage = Con.CloseRecordSet(RsEntetePage)
Set Myrange = Nothing
Set MySeet = Nothing
End Function
Function ExporteXlsComposants(Rs As Recordset, Id_IndiceProjet As Long)
Dim MySeet As EXCEL.Worksheet
Dim Myrange As EXCEL.Range
Dim Row As Long
Dim Rep As String
Dim NbColonne As Long
Dim TxtPoin As String
Dim NbLigne As Long
Dim Sql As String
Dim Fso As New FileSystemObject
Dim PathComposantsDefault As String
Dim RsComposants As Recordset
    NbLigne = 0
 Set MySeet = IsertSheet(MyWorkbook, "Composants", True)
 IsertColonne "ACTIVER", MySeet
DeleteRow MySeet

Set Myrange = MySeet.Range("A1").CurrentRegion
NbColonne = Myrange.Columns.Count
 If Rs.EOF = False Then

     Sql = "SELECT T_indiceProjet.Client FROM T_indiceProjet "
    Sql = Sql & "WHERE T_indiceProjet.Id=" & Rs!Id_IndiceProjet & ";"

    NumErr = 1
    
      Sql = "SELECT T_indiceProjet.Client FROM T_indiceProjet "
    Sql = Sql & "WHERE T_indiceProjet.Id=" & Rs!Id_IndiceProjet & ";"

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
                 PathComposantsDefault = RsComposants!PathComposants

End If
Else
 PathComposantsDefault = TableauPath.Item("PathComposantsDefault")
End If
If Left(PathComposantsDefault, 2) <> "\\" And Left(PathComposantsDefault, 1) = "\" Then PathComposantsDefault = TableauPath.Item("PathServer") & PathComposantsDefault
If Right(PathComposantsDefault, 2) = "\\" Then PathComposantsDefault = Mid(PathComposantsDefault, 1, Len(PathComposantsDefault) - 1)

    Dim fs, f, f1, s, sf
  Myrange(1, NbColonne).AutoFilter
'  MyExcel.Visible = True
    Set f = Fso.GetFolder(PathComposantsDefault) '\\10.30.0.5\\donnees d entreprise\Utilitaires\cablage\Librairies\COMPOSANTS\")
    Set sf = f.SubFolders
    For Each f1 In sf
       NbColonne = NbColonne + 1
    Myrange(1, NbColonne) = f1.Name
    Myrange(1, NbColonne).Interior.ColorIndex = 15
    Next
  
Myrange(1, NbColonne).AutoFilter

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
    If Rs.EOF = True Then Exit Function
While Rs.EOF = False
NbLigne = NbLigne + 1

Rs.MoveNext
Wend
Rs.MoveFirst


Row = 2
 FormBarGrah.ProgressBar1.Value = 0
 FormBarGrah.ProgressBar1.Max = NbLigne
 FormBarGrah.ProgressBar1Caption.Caption = " Exporter liste des Composants :"
 Set Myrange = MySeet.Range("A1").CurrentRegion

While Rs.EOF = False
 IncremanteBarGrah FormBarGrah
DoEvents
 Myrange(Row, 1) = Replace(Replace("" & Rs!ACTIVER, "Faux", 0), "Vrai", 1)
  Myrange(Row, 2) = "'" & Rs!DESIGNCOMP
  Myrange(Row, 3) = "'" & Rs!NUMCOMP
  Myrange(Row, 4) = "'" & Rs!REFCOMP
  For i = 5 To Myrange.Columns.Count
            If Myrange(1, i) = "" & Rs!Path Then Myrange(Row, i) = 1 Else Myrange(Row, i) = 0
        Next i
  
 
    Row = Row + 1
    Rs.MoveNext
Wend
Set Myrange = MySeet.Range("A1").CurrentRegion

Sql = "SELECT  T_indiceProjet.Ensemble, T_indiceProjet.Equipement,[Li] & '_' & [LI_Indice] AS Liste, "
Sql = Sql & "[PI] & '_' & [PI_Indice] AS Piece, T_indiceProjet.Client, T_indiceProjet.CleAc "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Id_IndiceProjet & ";"
Dim RsEntetePage As Recordset
Set RsEntetePage = Con.OpenRecordSet(Sql)

MiseEnPage MySeet, Myrange, "Affaire : " & RsEntetePage!CleAc & vbCrLf & _
 _
     "Pièce : " & RsEntetePage!Piece & vbCrLf & "Liste : " & RsEntetePage!Liste, vbCrLf & "Câblage : " & Replace("" & RsEntetePage!Ensemble, vbCrLf, " ") & vbCrLf & "Equipement : " & Replace("" & RsEntetePage!Equipement, vbCrLf, " ") & vbCrLf & _
     "" _
    , "ENCELADE" & vbCrLf & "Client : " & RsEntetePage!Client & vbCrLf & Format(Date, "dd-mmm-yyyy"), "", "&P/&N", "", 80, "C2", True, 2

MaJEncadreXls Myrange, xlThin, xlThin, xlHairline, xlHairline

Set RsEntetePage = Con.CloseRecordSet(RsEntetePage)
Set Myrange = Nothing
Set MySeet = Nothing
End Function
Function OpenModelXlt(Fichier As String) As EXCEL.Workbook
Set MyExcel = New EXCEL.Application
'MyExcel.Visible = True
Set OpenModelXlt = MyExcel.Workbooks.Open(Fichier)
End Function
Public Function IsertSheet(MyWorkbook As EXCEL.Workbook, Name As String, Optional Fin As Boolean) As EXCEL.Worksheet
On Error Resume Next
If Name = "Appro Connectique" Then
    Set IsertSheet = MyWorkbook.Sheets("Appro")
    If Err Then
        Err.Clear
        GoTo ReTest
    End If
Else
ReTest:
    Set IsertSheet = MyWorkbook.Sheets(Name)
    If Err Then
    Err.Clear
    If Fin = False Then
 Set IsertSheet = MyWorkbook.Sheets.Add(Before:=MyWorkbook.Sheets(1))
 Else
 Set IsertSheet = MyWorkbook.Sheets.Add(After:=MyWorkbook.Sheets(MyWorkbook.Sheets.Count))
End If

End If


' Before
 
 End If
 IsertSheet.Name = Name
 On Error GoTo 0
End Function

Public Sub FormatExcelPlage(Plage As Range, Couleur As Long, Merge As Boolean, Grille As Boolean, HorizontalAlignment As Long, VerticalAlignment As Long)
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
Function DeleteRow(MySheet As Worksheet, Optional Tous As Boolean = False)
Dim Myrange As Range
Set Myrange = MySheet.Range("A1").CurrentRegion
If Tous = True Then
MySheet.Cells.Delete Shift:=xlUp
Else
For i = Myrange.Rows.Count To 2 Step -1
    MySheet.Rows(CStr(i) & ":" & CStr(i)).Delete Shift:=xlUp
Next
End If
Set Myrange = Nothing
End Function
Public Sub EporteSynthese(Optional Affaire As Long)
Dim Sql As String
Dim Myrep As String
Dim Rs As Recordset
Dim MyExcel As New EXCEL.Application
Dim MyWorkbook As EXCEL.Workbook
Dim MySheet As Worksheet
Dim Fso As New FileSystemObject
Dim PathAffair As String
Dim Row As Long
Dim NameFichier
 Set TableauPath = funPath
'MyExcel.Visible = True
Set MyWorkbook = MyExcel.Workbooks.Add
Set MySheet = IsertSheet(MyWorkbook, "Synthèse", False)
Sql = "SELECT Rq_Synthese_Total.* FROM Rq_Synthese_Total;"
Set Rs = Con.OpenRecordSet(Sql)
If Affaire > 0 Then
    Rs.Filter = "Affaire=" & Affaire
    DoEvents
    PathAffair = "" & Rs!Client & "\PI\" & Rs!Affaire & "\"
End If
For i = 0 To Rs.Fields.Count - 1
MySheet.Cells(1, i + 1) = Rs.Fields(i).Name
Next
Row = 1
While Rs.EOF = False
Row = Row + 1
    For i = 0 To Rs.Fields.Count - 1
        MySheet.Cells(Row, i + 1) = Rs.Fields(i).Value
Next
Rs.MoveNext
Wend
NameFichier = TableauPath.Item("PathArchiveAutocad") '& "Synthèse.xls"
 If Left(NameFichier, 2) <> "\\" And Left(NameFichier, 1) = "\" Then NameFichier = TableauPath.Item("PathServer") & NameFichier & "\"
            If Right(NameFichier, 2) = "\\" Then NameFichier = Mid(NameFichier, 1, Len(NameFichier) - 1)
    If Right(NameFichier, 1) <> "\" Then NameFichier = NameFichier & "\"
If Affaire > 0 Then
    NameFichier = NameFichier & PathAffair
Else
    
   NameFichier = NameFichier
End If
NameFichier = NameFichier & "Synthèse"
If Affaire > 0 Then
    NameFichier = NameFichier & "_" & CStr(Affaire)
End If
NameFichier = NameFichier & ".XLS"
 If Fso.FileExists(NameFichier) = True Then
    Fso.DeleteFile NameFichier
    End If
   
    MiseEnPage MySheet, MySheet.Range("A1").CurrentRegion, "", "", "Date: " & Format(Date, "dd-mmm-yyyy"), "", "&P/&N", "", 80, "A2", True, 2, True
    
    MaJEncadreXls MySheet.Range("A1").CurrentRegion, xlThin, xlThin, xlHairline, xlHairline
    

    MyWorkbook.SaveAs NameFichier
    MyWorkbook.Close False
    MyExcel.Quit
Set MySheet = Nothing
Set MyWorkbook = Nothing
Set MyExcel = Nothing

End Sub
Public Sub MaJEncadreXls(Myrange As Range, LeftWeight As Long, RightWeight As Long, TopWeight As Long, BottomWeight As Long)
On Error Resume Next

'
' Macro3 Macro
' Macro enregistrée le 14/03/2005 par robert.durupt
'

'
    Myrange.Borders(xlDiagonalDown).LineStyle = xlNone
    Myrange.Borders(xlDiagonalUp).LineStyle = xlNone
    
    Myrange.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Myrange.Borders(xlEdgeLeft).Weight = LeftWeight
   
        Myrange.Borders(xlEdgeRight).LineStyle = xlContinuous
        Myrange.Borders(xlEdgeRight).Weight = RightWeight
        
      
        Myrange.Borders(xlEdgeTop).LineStyle = xlContinuous
        Myrange.Borders(xlEdgeTop).Weight = TopWeight
       
    
        Myrange.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Myrange.Borders(xlEdgeBottom).Weight = BottomWeight
      
    Myrange.Borders(xlInsideVertical).Weight = LeftWeight
     Myrange.Borders(xlInsideHorizontal).Weight = BottomWeight
 On Error GoTo 0
End Sub

Public Sub ExcelCreatTitre(MyRangeStrate As Range, RsTitre As Recordset, Optional Value As Boolean)
Dim Row As Long
Dim Col As Long
Col = MyRangeStrate.Column
For i = 0 To RsTitre.Fields.Count - 1
Debug.Print UCase(RsTitre.Fields(i).Name)
    If Value = False Then
    
       
        MyRangeStrate(1, Col + i) = UCase(RsTitre.Fields(i).Name)
       
    Else
        If RsTitre.Fields(i).Type = 11 Then
            If RsTitre.Fields(i).Value = True Then
                 MyRangeStrate(1, Col + i) = 1
                
            Else
                 MyRangeStrate(1, Col + i) = 0
            End If
        
        Else
        MyRangeStrate(1, Col + i) = "'" & RsTitre.Fields(i).Value
        End If
    End If
    
Next
End Sub
