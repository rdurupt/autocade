Attribute VB_Name = "ImporterXls"
Sub ImporteXls(Xls As String, IdIndiceProjet As Long, Optional Edition As Boolean)
Dim Sql As String
Dim Rs As Recordset
Dim IdProjet As Long
Dim IdIndice As Long
Dim IndexLigne As Long
DoEvents

Set TableauPath = funPath
IdIndice = IdIndiceProjet
Sql = "UPDATE T_indiceProjet SET T_indiceProjet.PlOk = False, T_indiceProjet.OuOk = False "
Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndice & ";"
Con.Execute Sql

Dim MyExcel As New EXCEL.Application
MyExcel.DisplayAlerts = False
'MyExcel.Visible = True
Dim MyClasseur As EXCEL.Workbook
Dim MySheet As EXCEL.Worksheet
Dim MyRange As EXCEL.Range
Set MyClasseur = MyExcel.Workbooks.Open(Xls)
'MyClasseur.Application.Visible = True
'***********************************************************************************************************************
'*                                          Supprime le contenu des tables temporaire.                                 *
Sql = "DELETE Xls_Critères.*  FROM Xls_Critères "
Sql = Sql & "where Xls_Critères.Job=" & NmJob & ";"
Con.Execute Sql


Sql = "DELETE xls_Ligne_Tableau_fils.*  FROM xls_Ligne_Tableau_fils "
Sql = Sql & "where xls_Ligne_Tableau_fils.Job=" & NmJob & ";"
Con.Execute Sql

Sql = "DELETE Xls_Connecteurs.* FROM Xls_Connecteurs "
Sql = Sql & "WHERE Xls_Connecteurs.Job=" & NmJob & ";"
Con.Execute Sql

Sql = "DELETE Xls_Nota.* FROM Xls_Nota "
Sql = Sql & "WHERE Xls_Nota.Job=" & NmJob & ";"
Con.Execute Sql

Sql = "DELETE Xls_Composants.* FROM Xls_Composants "
Sql = Sql & "WHERE Xls_Composants.Job=" & NmJob & ";"
Con.Execute Sql

Sql = "DELETE Xls_Noeuds.* FROM Xls_Noeuds "
Sql = Sql & "WHERE Xls_Noeuds.Job=" & NmJob & ";"
Con.Execute Sql
'MyExcel.Visible = True

'***********************************************************************************************************************
'*                                          Supprime le contenu des tables Ecart.                                     *
Sql = "DELETE T_Critères_Ecart.*  FROM T_Critères_Ecart "
Sql = Sql & "where T_Critères_Ecart.Id_IndiceProjet=" & IdIndiceProjet & ";"
Con.Execute Sql


Sql = "DELETE Ligne_Tableau_fils_Ecart.*  FROM Ligne_Tableau_fils_Ecart "
Sql = Sql & "where Ligne_Tableau_fils_Ecart.Id_IndiceProjet=" & IdIndiceProjet & ";"
Con.Execute Sql

Sql = "DELETE Connecteurs_Ecart.* FROM Connecteurs_Ecart "
Sql = Sql & "WHERE Connecteurs_Ecart.Id_IndiceProjet=" & IdIndiceProjet & ";"
Con.Execute Sql

Sql = "DELETE Nota_Ecart.* FROM Nota_Ecart "
Sql = Sql & "WHERE Nota_Ecart.Id_IndiceProjet=" & IdIndiceProjet & ";"
Con.Execute Sql

Sql = "DELETE Composants_Ecart.* FROM Composants_Ecart "
Sql = Sql & "WHERE Composants_Ecart.Id_IndiceProjet=" & IdIndiceProjet & ";"
Con.Execute Sql

Sql = "DELETE T_Noeuds_Ecart.* FROM T_Noeuds_Ecart "
Sql = Sql & "WHERE T_Noeuds_Ecart.Id_IndiceProjet=" & IdIndiceProjet & ";"
Con.Execute Sql

'***********************************************************************************************************************
'*                                        Sauvegarde les anciennes valeurs                                            *

Sql = "INSERT INTO T_Critères_Ecart SELECT T_Critères.* FROM T_Critères "
Sql = Sql & "WHERE T_Critères.Id_IndiceProjet=" & IdIndiceProjet & ";"
Con.Execute Sql

Sql = "INSERT INTO Connecteurs_Ecart SELECT Connecteurs.* FROM Connecteurs "
Sql = Sql & "WHERE Connecteurs.Id_IndiceProjet=" & IdIndiceProjet & ";"
Con.Execute Sql

Sql = "INSERT INTO Nota_Ecart SELECT Nota.* FROM Nota "
Sql = Sql & "WHERE Nota.Id_IndiceProjet=" & IdIndiceProjet & ";"
Con.Execute Sql

Sql = "INSERT INTO Composants_Ecart SELECT Composants.* FROM Composants "
Sql = Sql & "WHERE Composants.Id_IndiceProjet=" & IdIndiceProjet & ";"
Con.Execute Sql

Sql = "INSERT INTO Ligne_Tableau_fils_Ecart SELECT Ligne_Tableau_fils.* FROM Ligne_Tableau_fils "
Sql = Sql & "WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & IdIndiceProjet & ";"
Con.Execute Sql

Sql = "INSERT INTO T_Noeuds_Ecart SELECT T_Noeuds.* FROM T_Noeuds "
Sql = Sql & "WHERE T_Noeuds.Id_IndiceProjet=" & IdIndiceProjet & ";"
Con.Execute Sql

'***********************************************************************************************************************
'*                              Importe la liste des Noeuds dans la table temporaire:                                    *

Set MySheet = MyClasseur.Worksheets("Noeuds")
MySheet.Activate
'MySheet.Application.Visible = True
Set MyRange = MySheet.Range("A1").CurrentRegion
 FormBarGrah.ProgressBar1.Value = 0
 FormBarGrah.ProgressBar1.Max = MyRange.Rows.Count
 FormBarGrah.ProgressBar1Caption.Caption = " Importe la liste des Noeuds "

For IndexLigne = 2 To MyRange.Rows.Count
 IncremanteBarGrah FormBarGrah
DoEvents

   Sql = sqlRange(MyRange, IndexLigne, "Xls_Noeuds")
   Con.Execute Sql
Next IndexLigne

'***********************************************************************************************************************
'*                              Importe la liste des Critères dans la table temporaire:                                    *

Set MySheet = MyClasseur.Worksheets("Critères")

Set MyRange = MySheet.Range("A1").CurrentRegion
 FormBarGrah.ProgressBar1.Value = 0
 FormBarGrah.ProgressBar1.Max = MyRange.Rows.Count
 FormBarGrah.ProgressBar1Caption.Caption = " Importe la liste de Critères "

For Row = 2 To MyRange.Rows.Count
 IncremanteBarGrah FormBarGrah
DoEvents

   Sql = sqlRange(MyRange, IndexLigne, "Xls_Critères")
   Con.Execute Sql
Next Row

'***********************************************************************************************************************
'*                              Importe la liste des fils dans la table temporaire:                                    *

Set MySheet = MyClasseur.Worksheets("Ligne_Tableau_fils")
MySheet.Activate
Set MyRange = MySheet.Range("A1").CurrentRegion
 FormBarGrah.ProgressBar1.Value = 0
 FormBarGrah.ProgressBar1.Max = MyRange.Rows.Count
 FormBarGrah.ProgressBar1Caption.Caption = " Importe la liste de fils"

For Row = 2 To MyRange.Rows.Count
 IncremanteBarGrah FormBarGrah
DoEvents

   Sql = sqlRange(MyRange, IndexLigne, "Xls_Ligne_Tableau_fils")
   Con.Execute Sql
Next Row

'***********************************************************************************************************************
'*                              Importe la liste des Connecteurs dans la table temporaire:                             *

Set MySheet = MyClasseur.Worksheets("Connecteurs")
MySheet.Activate
Set MyRange = MySheet.Range("A1").CurrentRegion
 
 FormBarGrah.ProgressBar1.Value = 0
 FormBarGrah.ProgressBar1.Max = MyRange.Rows.Count
 FormBarGrah.ProgressBar1Caption.Caption = " Importe la liste des Connecteurs"

For Row = 2 To MyRange.Rows.Count
 IncremanteBarGrah FormBarGrah
   Sql = sqlRange(MyRange, IndexLigne, "Xls_Connecteurs")
   Con.Execute Sql
Next Row

'***********************************************************************************************************************
'*                              Importe la liste des Composants dans la table temporaire:                              *

Set MySheet = MyClasseur.Worksheets("Composants")
MySheet.Activate
Set MyRange = MySheet.Range("A1").CurrentRegion

 FormBarGrah.ProgressBar1.Value = 0
 FormBarGrah.ProgressBar1.Max = MyRange.Rows.Count
 FormBarGrah.ProgressBar1Caption.Caption = " Importe la liste des Composants"

For Row = 2 To MyRange.Rows.Count
 IncremanteBarGrah FormBarGrah
   Sql = sqlRange(MyRange, IndexLigne, "Xls_Composants")
   Con.Execute Sql
Next Row

'***********************************************************************************************************************
'*                              Importe la liste des Notas dans la table temporaire:                                   *

Set MySheet = MyClasseur.Worksheets("Notas")
MySheet.Activate
Set MyRange = MySheet.Range("A1").CurrentRegion
 
 FormBarGrah.ProgressBar1.Value = 0
 FormBarGrah.ProgressBar1.Max = MyRange.Rows.Count
 FormBarGrah.ProgressBar1Caption.Caption = " Importe la liste des Notas"

For Row = 2 To MyRange.Rows.Count
 IncremanteBarGrah FormBarGrah
   Sql = sqlRange(MyRange, IndexLigne, "Xls_Nota")
   Con.Execute Sql
Next Row

Set MyRange = Nothing
Set MySheet = Nothing
'***********************************************************************************************************************
'*                      si ImportteXls est en mod édition
If Edition = True Then
Set MySheet = IsertSheet(MyClasseur, "Nomenclature Habillage", True)
MySheet.Activate
    insertExelAccess MySheet, "T_Appro_Habillage", 1, IdIndiceProjet
Set MySheet = IsertSheet(MyClasseur, "Nomenclature Fils", True)
    insertExelAccess MySheet, "T_Prix_Fils", 1, IdIndiceProjet
Set MySheet = IsertSheet(MyClasseur, "Nomenclature Connecteur", True)
    insertExelAccess MySheet, "T_Nomenclature", 1, IdIndiceProjet
End If
MyClasseur.Close False
Set MyClasseur = Nothing
MyExcel.Quit
Set MyExcel = Nothing


MajBase IdIndice
MajEcart IdIndiceProjet, MyExcel

'***********************************************************************************************************************
'*                                          Supprime le contenu des tables Ecart.                                     *
Sql = "DELETE T_Critères_Ecart.*  FROM T_Critères_Ecart "
Sql = Sql & "where T_Critères_Ecart.Id_IndiceProjet=" & IdIndiceProjet & ";"
Con.Execute Sql


Sql = "DELETE Ligne_Tableau_fils_Ecart.*  FROM Ligne_Tableau_fils_Ecart "
Sql = Sql & "where Ligne_Tableau_fils_Ecart.Id_IndiceProjet=" & IdIndiceProjet & ";"
Con.Execute Sql

Sql = "DELETE Connecteurs_Ecart.* FROM Connecteurs_Ecart "
Sql = Sql & "WHERE Connecteurs_Ecart.Id_IndiceProjet=" & IdIndiceProjet & ";"
Con.Execute Sql

Sql = "DELETE Nota_Ecart.* FROM Nota_Ecart "
Sql = Sql & "WHERE Nota_Ecart.Id_IndiceProjet=" & IdIndiceProjet & ";"
Con.Execute Sql

Sql = "DELETE Composants_Ecart.* FROM Composants_Ecart "
Sql = Sql & "WHERE Composants_Ecart.Id_IndiceProjet=" & IdIndiceProjet & ";"
Con.Execute Sql

Sql = "DELETE T_Noeuds_Ecart.* FROM T_Noeuds_Ecart "
Sql = Sql & "WHERE T_Noeuds_Ecart.Id_IndiceProjet=" & IdIndiceProjet & ";"
Con.Execute Sql

'***********************************************************************************************************************
'*                                          Supprime le contenu des tables temporaire.                                 *

Sql = "DELETE xls_Ligne_Tableau_fils.*  FROM xls_Ligne_Tableau_fils "
Sql = Sql & "where xls_Ligne_Tableau_fils.Job=" & NmJob & ";"
Con.Execute Sql

Sql = "DELETE Xls_Connecteurs.* FROM Xls_Connecteurs "
Sql = Sql & "WHERE Xls_Connecteurs.Job=" & NmJob & ";"
Con.Execute Sql

Sql = "DELETE Xls_Nota.* FROM Xls_Nota "
Sql = Sql & "WHERE Xls_Nota.Job=" & NmJob & ";"
Con.Execute Sql

Sql = "DELETE Xls_Composants.* FROM Xls_Composants "
Sql = Sql & "WHERE Xls_Composants.Job=" & NmJob & ";"
Con.Execute Sql

 FormBarGrah.ProgressBar1.Value = 0
 '***********************************************************************************************************************

 FormBarGrah.ProgressBar1Caption.Caption = " Fin du traitement"
End Sub
Sub Importefrm(FRM As Object, Optional Edition As Boolean)
Dim Sql As String
Dim Rs As Recordset
Dim IdProjet As Long
Dim IdIndice As Long
Dim BarMax As Long
Dim IndexLigne As Long
DoEvents

Set TableauPath = funPath
IdIndice = FRM.Tag
Sql = "UPDATE T_indiceProjet SET T_indiceProjet.PlOk = False, T_indiceProjet.OuOk = False "
Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndice & ";"
Con.Execute Sql

'Dim MyExcel As New EXCEL.Application
'MyExcel.DisplayAlerts = False
'MyExcel.Visible = True
'Dim MyClasseur As EXCEL.Workbook
'Dim MySheet As EXCEL.Worksheet
Dim MyRange As Object
'Set MyClasseur = MyExcel.Workbooks.Open(Xls)
'MyClasseur.Application.Visible = True
'***********************************************************************************************************************
'*                                          Supprime le contenu des tables temporaire.                                 *
Sql = "DELETE Xls_Critères.*  FROM Xls_Critères "
Sql = Sql & "where Xls_Critères.Job=" & NmJob & ";"
Con.Execute Sql


Sql = "DELETE   FROM xls_Ligne_Tableau_fils "
Sql = Sql & "where xls_Ligne_Tableau_fils.Job=" & NmJob & ";"
Con.Execute Sql

Sql = "DELETE  FROM Xls_Connecteurs "
Sql = Sql & "WHERE Xls_Connecteurs.Job=" & NmJob & ";"
Con.Execute Sql

Sql = "DELETE  FROM Xls_Nota "
Sql = Sql & "WHERE Xls_Nota.Job=" & NmJob & ";"
Con.Execute Sql

Sql = "DELETE  FROM Xls_Composants "
Sql = Sql & "WHERE Xls_Composants.Job=" & NmJob & ";"
Con.Execute Sql

Sql = "DELETE  FROM Xls_Noeuds "
Sql = Sql & "WHERE Xls_Noeuds.Job=" & NmJob & ";"
Con.Execute Sql
'MyExcel.Visible = True

'***********************************************************************************************************************
'*                                          Supprime le contenu des tables Ecart.                                     *
Sql = "DELETE   FROM T_Critères_Ecart "
Sql = Sql & "where T_Critères_Ecart.Id_IndiceProjet=" & FRM.Tag & ";"
Con.Execute Sql


Sql = "DELETE   FROM Ligne_Tableau_fils_Ecart "
Sql = Sql & "where Ligne_Tableau_fils_Ecart.Id_IndiceProjet=" & FRM.Tag & ";"
Con.Execute Sql

Sql = "DELETE  FROM Connecteurs_Ecart "
Sql = Sql & "WHERE Connecteurs_Ecart.Id_IndiceProjet=" & FRM.Tag & ";"
Con.Execute Sql

Sql = "DELETE  FROM Nota_Ecart "
Sql = Sql & "WHERE Nota_Ecart.Id_IndiceProjet=" & FRM.Tag & ";"
Con.Execute Sql

Sql = "DELETE  FROM Composants_Ecart "
Sql = Sql & "WHERE Composants_Ecart.Id_IndiceProjet=" & FRM.Tag & ";"
Con.Execute Sql

Sql = "DELETE  FROM T_Noeuds_Ecart "
Sql = Sql & "WHERE T_Noeuds_Ecart.Id_IndiceProjet=" & FRM.Tag & ";"
Con.Execute Sql

'***********************************************************************************************************************
'*                                        Sauvegarde les anciennes valeurs                                            *

Sql = "INSERT INTO T_Critères_Ecart SELECT T_Critères.* FROM T_Critères "
Sql = Sql & "WHERE T_Critères.Id_IndiceProjet=" & FRM.Tag & ";"
Con.Execute Sql

Sql = "INSERT INTO Connecteurs_Ecart SELECT Connecteurs.* FROM Connecteurs "
Sql = Sql & "WHERE Connecteurs.Id_IndiceProjet=" & FRM.Tag & ";"
Con.Execute Sql

Sql = "INSERT INTO Nota_Ecart SELECT Nota.* FROM Nota "
Sql = Sql & "WHERE Nota.Id_IndiceProjet=" & FRM.Tag & ";"
Con.Execute Sql

Sql = "INSERT INTO Composants_Ecart SELECT Composants.* FROM Composants "
Sql = Sql & "WHERE Composants.Id_IndiceProjet=" & FRM.Tag & ";"
Con.Execute Sql

Sql = "INSERT INTO Ligne_Tableau_fils_Ecart SELECT Ligne_Tableau_fils.* FROM Ligne_Tableau_fils "
Sql = Sql & "WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & FRM.Tag & ";"
Con.Execute Sql

Sql = "INSERT INTO T_Noeuds_Ecart SELECT T_Noeuds.* FROM T_Noeuds "
Sql = Sql & "WHERE T_Noeuds.Id_IndiceProjet=" & FRM.Tag & ";"
Con.Execute Sql

'***********************************************************************************************************************
'*                              Importe la liste des Noeuds dans la table temporaire:                                    *
'
'Set MySheet = MyClasseur.Worksheets("Noeuds")
'MySheet.Activate
'MySheet.Application.Visible = True
Set MyRange = FRM.Noeuds.Range("A1").CurrentRegion
 FRM.ProgressBar1.Value = 0
 FRM.ProgressBar1.Max = RetourneNbRows(MyRange)
 FRM.ProgressBar1Caption.Caption = " Importe la liste des Noeuds "

For IndexLigne = 2 To MyRange.Rows.Count
 IncremanteBarGrah FRM

Sql = ""
   Sql = sqlRange(MyRange, IndexLigne, "Xls_Noeuds")
   DoEvents
   
   Con.Execute Sql
Next IndexLigne

'***********************************************************************************************************************
'*                              Importe la liste des Critères dans la table temporaire:                                    *

Set MyRange = FRM.Crit.Range("A1").CurrentRegion
 FRM.ProgressBar1.Value = 0
 FRM.ProgressBar1.Max = RetourneNbRows(MyRange)
 FRM.ProgressBar1Caption.Caption = " Importe la liste de Critères "

For IndexLigne = 2 To MyRange.Rows.Count
 IncremanteBarGrah FRM
DoEvents

   Sql = sqlRange(MyRange, IndexLigne, "Xls_Critères", "ID")
   Con.Execute Sql
Next IndexLigne

'***********************************************************************************************************************
'*                              Importe la liste des fils dans la table temporaire:                                    *

''Set MySheet = MyClasseur.Worksheets("Ligne_Tableau_fils")
'MySheet.Activate
Set MyRange = FRM.Fil.Range("A1").CurrentRegion
 FRM.ProgressBar1.Value = 0
 FRM.ProgressBar1.Max = RetourneNbRows(MyRange)
 FRM.ProgressBar1Caption.Caption = " Importe la liste de fils"
'FRM.NoMacro2 = True
For IndexLigne = 2 To MyRange.Rows.Count
'If IndexLigne = 4 Then MsgBox ""
 IncremanteBarGrah FRM
DoEvents

Sql = ""
   Sql = sqlRange(MyRange, IndexLigne, "Xls_Ligne_Tableau_fils")
   Debug.Print IndexLigne
  
   Con.Execute Sql
Next IndexLigne

'***********************************************************************************************************************
'*                              Importe la liste des Connecteurs dans la table temporaire:                             *

'Set MySheet = MyClasseur.Worksheets("Connecteurs")
'MySheet.Activate
Set MyRange = FRM.Conn.Range("A1").CurrentRegion
 
 FRM.ProgressBar1.Value = 0
 FRM.ProgressBar1.Max = RetourneNbRows(MyRange)
 FRM.ProgressBar1Caption.Caption = " Importe la liste des Connecteurs"

For IndexLigne = 2 To MyRange.Rows.Count
 IncremanteBarGrah FRM
   Sql = sqlRange(MyRange, IndexLigne, "Xls_Connecteurs")
   Con.Execute Sql
Next IndexLigne

'***********************************************************************************************************************
'*                              Importe la liste des Composants dans la table temporaire:                              *

'Set MySheet = MyClasseur.Worksheets("Composants")
'MySheet.Activate
Set MyRange = FRM.Comp.Range("A1").CurrentRegion

 FRM.ProgressBar1.Value = 0
 FRM.ProgressBar1.Max = RetourneNbRows(MyRange)
 FRM.ProgressBar1Caption.Caption = " Importe la liste des Composants"

For IndexLigne = 2 To MyRange.Rows.Count
 IncremanteBarGrah FRM
   Sql = sqlRange(MyRange, IndexLigne, "Xls_Composants")
   Con.Execute Sql
Next IndexLigne

'***********************************************************************************************************************
'*                              Importe la liste des Notas dans la table temporaire:                                   *

'Set MySheet = MyClasseur.Worksheets("Notas")
'MySheet.Activate
Set MyRange = FRM.Notas.Range("A1").CurrentRegion
 
 FRM.ProgressBar1.Value = 0
 BarMax = MyRange.Rows.Count - 1
 
 FRM.ProgressBar1.Max = RetourneNbRows(MyRange)
 FRM.ProgressBar1Caption.Caption = " Importe la liste des Notas"

For IndexLigne = 2 To MyRange.Rows.Count
 IncremanteBarGrah FRM
   Sql = sqlRange(MyRange, IndexLigne, "Xls_Nota")
   Con.Execute Sql
Next IndexLigne

Set MyRange = Nothing

'
MajBase IdIndice
Dim MyExcel As EXCEL.Application
Set MyExcel = New EXCEL.Application
MyExcel.DisplayAlerts = False
FRM.ProgressBar1.Value = 0
 FRM.ProgressBar1.Max = 1
 FRM.ProgressBar1Caption.Caption = " Mise ajour du fichier d'écarts."
MajEcart FRM.Tag, MyExcel

'***********************************************************************************************************************
'*                                          Supprime le contenu des tables Ecart.                                     *
Sql = "DELETE T_Critères_Ecart.*  FROM T_Critères_Ecart "
Sql = Sql & "where T_Critères_Ecart.Id_IndiceProjet=" & FRM.Tag & ";"
Con.Execute Sql


Sql = "DELETE Ligne_Tableau_fils_Ecart.*  FROM Ligne_Tableau_fils_Ecart "
Sql = Sql & "where Ligne_Tableau_fils_Ecart.Id_IndiceProjet=" & FRM.Tag & ";"
Con.Execute Sql

Sql = "DELETE Connecteurs_Ecart.* FROM Connecteurs_Ecart "
Sql = Sql & "WHERE Connecteurs_Ecart.Id_IndiceProjet=" & FRM.Tag & ";"
Con.Execute Sql

Sql = "DELETE Nota_Ecart.* FROM Nota_Ecart "
Sql = Sql & "WHERE Nota_Ecart.Id_IndiceProjet=" & FRM.Tag & ";"
Con.Execute Sql

Sql = "DELETE Composants_Ecart.* FROM Composants_Ecart "
Sql = Sql & "WHERE Composants_Ecart.Id_IndiceProjet=" & FRM.Tag & ";"
Con.Execute Sql

Sql = "DELETE T_Noeuds_Ecart.* FROM T_Noeuds_Ecart "
Sql = Sql & "WHERE T_Noeuds_Ecart.Id_IndiceProjet=" & FRM.Tag & ";"
Con.Execute Sql

'***********************************************************************************************************************
'*                                          Supprime le contenu des tables temporaire.                                 *

Sql = "DELETE xls_Ligne_Tableau_fils.*  FROM xls_Ligne_Tableau_fils "
Sql = Sql & "where xls_Ligne_Tableau_fils.Job=" & NmJob & ";"
Con.Execute Sql

Sql = "DELETE Xls_Connecteurs.* FROM Xls_Connecteurs "
Sql = Sql & "WHERE Xls_Connecteurs.Job=" & NmJob & ";"
Con.Execute Sql

Sql = "DELETE Xls_Nota.* FROM Xls_Nota "
Sql = Sql & "WHERE Xls_Nota.Job=" & NmJob & ";"
Con.Execute Sql

Sql = "DELETE Xls_Composants.* FROM Xls_Composants "
Sql = Sql & "WHERE Xls_Composants.Job=" & NmJob & ";"
Con.Execute Sql

 FRM.ProgressBar1.Value = 0
 '***********************************************************************************************************************

 FRM.ProgressBar1Caption.Caption = " Fin du traitement"
End Sub


Sub MajEcart(IdIndiceProjet As Long, MyExcel As EXCEL.Application)
 Set TableauPath = funPath
Dim L As Long
Dim C As Long
Dim boolSave As Boolean
Dim Sql As String
Dim RsSuprimer As Recordset
Dim RsAjouter As Recordset
Dim RsModifier As Recordset
Dim PathArchiveAutocad As String

boolSave = False
   PathArchiveAutocad = TableauPath.Item("PathArchiveAutocad")
 PathArchiveAutocad = DefinirChemienComplet(TableauPath.Item("PathServer"), PathArchiveAutocad)
Set MyWorkbook = MyExcel.Workbooks.Add
Set MyWorkbook = MyExcel.Workbooks.Add
For I = MyWorkbook.Worksheets.Count To 1 Step -1
    DeletSheet MyWorkbook.Worksheets(I)
Next
Sql = "SELECT Nota_Ecart.ACTIVER,Nota_Ecart.NOTA, Nota_Ecart.NUMNOTA ,Nota_Ecart.[OPTION],Nota_Ecart.COMMENTAIRES "
Sql = Sql & "FROM Nota_Ecart LEFT JOIN Nota ON Nota_Ecart.Id = Nota.Id "
Sql = Sql & "WHERE Nota_Ecart.Id_IndiceProjet=" & IdIndiceProjet & " "
Sql = Sql & "AND Nota.id Is Null;"



Set RsSuprimer = Con.OpenRecordSet(Sql)

Sql = "SELECT Nota.ACTIVER,Nota.NOTA, Nota.NUMNOTA ,Nota.[OPTION],Nota_Ecart.COMMENTAIRES "
Sql = Sql & "FROM Nota LEFT JOIN Nota_Ecart ON Nota.Id = Nota_Ecart.Id  "
Sql = Sql & "WHERE Nota.Id_IndiceProjet=" & IdIndiceProjet & "  "
Sql = Sql & "AND Nota_Ecart.id Is Null;"


Set RsAjouter = Con.OpenRecordSet(Sql)


Sql = "SELECT '' AS [Avant/Après],Nota_Ecart.ACTIVER, Nota_Ecart.NOTA, Nota_Ecart.NUMNOTA,Nota_Ecart.[OPTION],Nota_Ecart.COMMENTAIRES , '' AS Expr1,Nota.ACTIVER, Nota.NOTA, Nota.NUMNOTA,Nota.[OPTION],Nota.COMMENTAIRES "
Sql = Sql & "FROM Nota INNER JOIN Nota_Ecart ON Nota.Id = Nota_Ecart.Id  "
Sql = Sql & "WHERE Nota_Ecart.Id_IndiceProjet=" & IdIndiceProjet & ";  "


Set RsModifier = Con.OpenRecordSet(Sql)



L = 0: C = 0
If MajEcartExcel(IdIndiceProjet, MyWorkbook, L, C, RsSuprimer, RsAjouter, RsModifier, "Notas_Ecart") = True Then boolSave = True

Sql = "SELECT Composants_Ecart.ACTIVER, Composants_Ecart.DESIGNCOMP, Composants_Ecart.NUMCOMP, Composants_Ecart.REFCOMP, Composants_Ecart.Path,Composants_Ecart.[OPTION],Composants_Ecart.COMMENTAIRES "
Sql = Sql & ",Composants_Ecart.Code_APP_Lier,Composants_Ecart.Voie,Composants_Ecart.POS,Composants_Ecart.[POS-OUT] "
Sql = Sql & "FROM Composants_Ecart LEFT JOIN Composants ON Composants_Ecart.Id = Composants.Id "
Sql = Sql & "WHERE Composants.id Is Null "
Sql = Sql & "AND Composants_Ecart.Id_IndiceProjet=" & IdIndiceProjet & " ;"


Set RsSuprimer = Con.OpenRecordSet(Sql)

Sql = "SELECT  Composants.ACTIVER,Composants.DESIGNCOMP, Composants.NUMCOMP, Composants.REFCOMP, Composants.Path,Composants.[OPTION] ,Composants.COMMENTAIRES "
Sql = Sql & ",Composants.Code_APP_Lier,Composants.Voie,Composants.POS,Composants.[POS-OUT] "
Sql = Sql & "FROM Composants LEFT JOIN Composants_Ecart ON Composants.Id= Composants_Ecart.Id "
Sql = Sql & "WHERE Composants.Id_IndiceProjet=" & IdIndiceProjet & "  "
Sql = Sql & "AND Composants_Ecart.id Is Null;"

Set RsAjouter = Con.OpenRecordSet(Sql)

Sql = "SELECT '' AS [Avant/Après],Composants_Ecart.ACTIVER, Composants_Ecart.DESIGNCOMP, Composants_Ecart.NUMCOMP,   "
Sql = Sql & "Composants_Ecart.REFCOMP, Composants_Ecart.Path,Composants_Ecart.[OPTION],Composants_Ecart.COMMENTAIRES "
Sql = Sql & ",Composants_Ecart.Code_APP_Lier,Composants_Ecart.Voie,Composants_Ecart.POS,Composants_Ecart.[POS-OUT] "
Sql = Sql & ", '' AS Expr1,Composants.ACTIVER, Composants.DESIGNCOMP,   "
Sql = Sql & "Composants.NUMCOMP, Composants.REFCOMP, Composants.Path,Composants.[OPTION],Composants.COMMENTAIRES "
Sql = Sql & ",Composants.Code_APP_Lier,Composants.Voie,Composants.POS,Composants.[POS-OUT] "
Sql = Sql & "FROM Composants_Ecart INNER JOIN Composants ON Composants_Ecart.Id = Composants.Id "
Sql = Sql & "WHERE Composants.Id_IndiceProjet=" & IdIndiceProjet & "  "


Set RsModifier = Con.OpenRecordSet(Sql)

L = 0: C = 0
If MajEcartExcel(IdIndiceProjet, MyWorkbook, L, C, RsSuprimer, RsAjouter, RsModifier, "Composants_Ecart") = True Then boolSave = True




Sql = "SELECT  T_Noeuds_Ecart.ACTIVER,T_Noeuds_Ecart.Fleche_Droite,T_Noeuds_Ecart.TORON_PRINCIPAL,  "
Sql = Sql & "T_Noeuds_Ecart.NŒUDS, T_Noeuds_Ecart.LONGUEUR, T_Noeuds_Ecart.LONGUEUR_CUMULEE, "
Sql = Sql & "T_Noeuds_Ecart.DESIGN_HAB, T_Noeuds_Ecart.CODE_RSA, T_Noeuds_Ecart.CODE_PSA, "
Sql = Sql & "T_Noeuds_Ecart.CODE_ENC, T_Noeuds_Ecart.DIAMETRE, T_Noeuds_Ecart.CLASSE_T,T_Noeuds_Ecart.[OPTION],T_Noeuds_Ecart.COMMENTAIRES "
Sql = Sql & "FROM T_Noeuds_Ecart LEFT JOIN T_Noeuds "
Sql = Sql & "ON T_Noeuds_Ecart.Id = T_Noeuds.Id "
Sql = Sql & "WHERE T_Noeuds.id Is Null "
Sql = Sql & "AND T_Noeuds_Ecart.Id_IndiceProjet=" & IdIndiceProjet & " ;"


Set RsSuprimer = Con.OpenRecordSet(Sql)

Sql = "SELECT T_Noeuds.ACTIVER, T_Noeuds.Fleche_Droite, T_Noeuds.TORON_PRINCIPAL,  "
Sql = Sql & "T_Noeuds.NŒUDS, T_Noeuds.LONGUEUR, T_Noeuds.LONGUEUR_CUMULEE, T_Noeuds.DESIGN_HAB,  "
Sql = Sql & "T_Noeuds.CODE_RSA, T_Noeuds.CODE_PSA, T_Noeuds.CODE_ENC, T_Noeuds.DIAMETRE, T_Noeuds.CLASSE_T,T_Noeuds.[OPTION],T_Noeuds.COMMENTAIRES "
Sql = Sql & "FROM T_Noeuds LEFT JOIN T_Noeuds_Ecart  "
Sql = Sql & "ON T_Noeuds.Id = T_Noeuds_Ecart.Id  "
Sql = Sql & "WHERE T_Noeuds.Id_IndiceProjet=" & IdIndiceProjet & "  "
Sql = Sql & "AND T_Noeuds_Ecart.id Is Null;"

Set RsAjouter = Con.OpenRecordSet(Sql)

Sql = "SELECT '' AS [Avant/Après],T_Noeuds_Ecart.ACTIVER, T_Noeuds_Ecart.Fleche_Droite, T_Noeuds_Ecart.TORON_PRINCIPAL,  "
Sql = Sql & " T_Noeuds_Ecart.NŒUDS, T_Noeuds_Ecart.LONGUEUR,  "
Sql = Sql & "T_Noeuds_Ecart.LONGUEUR_CUMULEE, T_Noeuds_Ecart.DESIGN_HAB, T_Noeuds_Ecart.CODE_RSA,  "
Sql = Sql & "T_Noeuds_Ecart.CODE_PSA, T_Noeuds_Ecart.CODE_ENC, T_Noeuds_Ecart.DIAMETRE,  "
Sql = Sql & "T_Noeuds_Ecart.CLASSE_T,T_Noeuds_Ecart.[OPTION],T_Noeuds_Ecart.COMMENTAIRES , '' AS Expr1,T_Noeuds.ACTIVER, T_Noeuds.Fleche_Droite, T_Noeuds.TORON_PRINCIPAL,  "
Sql = Sql & " T_Noeuds.NŒUDS, T_Noeuds.LONGUEUR, T_Noeuds.LONGUEUR_CUMULEE,  "
Sql = Sql & "T_Noeuds.DESIGN_HAB, T_Noeuds.CODE_RSA, T_Noeuds.CODE_PSA, T_Noeuds.CODE_ENC,  "
Sql = Sql & "T_Noeuds.DIAMETRE, T_Noeuds.CLASSE_T,T_Noeuds.[OPTION],T_Noeuds.COMMENTAIRES "
Sql = Sql & "FROM T_Noeuds INNER JOIN T_Noeuds_Ecart  "
Sql = Sql & "ON T_Noeuds.Id = T_Noeuds_Ecart.Id "
Sql = Sql & "WHERE T_Noeuds_Ecart.Id_IndiceProjet=" & IdIndiceProjet & "  "


Set RsModifier = Con.OpenRecordSet(Sql)
L = 0: C = 0
If MajEcartExcel(IdIndiceProjet, MyWorkbook, L, C, RsSuprimer, RsAjouter, RsModifier, "Noeuds_Ecart") = True Then boolSave = True

Sql = "SELECT Ligne_Tableau_fils_Ecart.LIAI, Ligne_Tableau_fils_Ecart.DESIGNATION, Ligne_Tableau_fils_Ecart.FIL, Ligne_Tableau_fils_Ecart.SECT, "
Sql = Sql & "Ligne_Tableau_fils_Ecart.TEINT , Ligne_Tableau_fils_Ecart.TEINT2, Ligne_Tableau_fils_Ecart.ISO, Ligne_Tableau_fils_Ecart.Long,  "
Sql = Sql & "Ligne_Tableau_fils_Ecart.[LONG CP], Ligne_Tableau_fils_Ecart.Coupe, Ligne_Tableau_fils_Ecart.POS, Ligne_Tableau_fils_Ecart.[POS-OUT],  "
Sql = Sql & "Ligne_Tableau_fils_Ecart.FA, Ligne_Tableau_fils_Ecart.App, Ligne_Tableau_fils_Ecart.VOI, Ligne_Tableau_fils_Ecart.[Ref Connecteur],  "
Sql = Sql & "Ligne_Tableau_fils_Ecart.[Ref Connecteur_Four], Ligne_Tableau_fils_Ecart.Long_Add, Ligne_Tableau_fils_Ecart.[Ref Clip],  "
Sql = Sql & "Ligne_Tableau_fils_Ecart.[Ref Clip Four], Ligne_Tableau_fils_Ecart.[Ref Joint], Ligne_Tableau_fils_Ecart.[Ref Joint four],  "
Sql = Sql & "  Ligne_Tableau_fils_Ecart.POS2,  "
Sql = Sql & "Ligne_Tableau_fils_Ecart.[POS-OUT2], Ligne_Tableau_fils_Ecart.FA2, Ligne_Tableau_fils_Ecart.APP2, Ligne_Tableau_fils_Ecart.VOI2,  "
Sql = Sql & "Ligne_Tableau_fils_Ecart.[Ref Connecteur2] , Ligne_Tableau_fils_Ecart.[Ref Connecteur_Four2],  "
Sql = Sql & "Ligne_Tableau_fils_Ecart.Long_Add2, Ligne_Tableau_fils_Ecart.[Ref Clip2], Ligne_Tableau_fils_Ecart.[Ref Clip Four2],  "
Sql = Sql & "Ligne_Tableau_fils_Ecart.[Ref Joint2], Ligne_Tableau_fils_Ecart.[Ref Joint four2],   "
Sql = Sql & " Ligne_Tableau_fils_Ecart.PRECO,  "
Sql = Sql & "Ligne_Tableau_fils_Ecart.PRECO2, Ligne_Tableau_fils_Ecart.PRECOG, Ligne_Tableau_fils_Ecart.Option, Ligne_Tableau_fils_Ecart.Activer,  "
Sql = Sql & "Ligne_Tableau_fils_Ecart.[Critères spécifiques],Ligne_Tableau_fils_Ecart.COMMENTAIRES "


Sql = Sql & "FROM Ligne_Tableau_fils_Ecart LEFT JOIN Ligne_Tableau_fils  "
Sql = Sql & "ON Ligne_Tableau_fils_Ecart.Id = Ligne_Tableau_fils.Id  "
Sql = Sql & "WHERE Ligne_Tableau_fils.id Is Null AND Ligne_Tableau_fils_Ecart.Id_IndiceProjet=" & IdIndiceProjet & ";  "







Set RsSuprimer = Con.OpenRecordSet(Sql)



Sql = "SELECT Ligne_Tableau_fils.LIAI, Ligne_Tableau_fils.DESIGNATION, Ligne_Tableau_fils.FIL, Ligne_Tableau_fils.SECT, Ligne_Tableau_fils.TEINT, "
Sql = Sql & "Ligne_Tableau_fils.TEINT2 , Ligne_Tableau_fils.ISO, Ligne_Tableau_fils.Long, Ligne_Tableau_fils.[LONG CP], Ligne_Tableau_fils.Coupe,  "
Sql = Sql & "Ligne_Tableau_fils.POS, Ligne_Tableau_fils.[POS-OUT], Ligne_Tableau_fils.FA, Ligne_Tableau_fils.App, Ligne_Tableau_fils.VOI,  "
Sql = Sql & "Ligne_Tableau_fils.[Ref Connecteur], Ligne_Tableau_fils.[Ref Connecteur_Four], Ligne_Tableau_fils.Long_Add,  "
Sql = Sql & "Ligne_Tableau_fils.[Ref Clip], Ligne_Tableau_fils.[Ref Clip Four], Ligne_Tableau_fils.[Ref Joint],  "
Sql = Sql & "Ligne_Tableau_fils.[Ref Joint four],   "
Sql = Sql & "Ligne_Tableau_fils.POS2, Ligne_Tableau_fils.[POS-OUT2], Ligne_Tableau_fils.FA2, Ligne_Tableau_fils.APP2,  "
Sql = Sql & "Ligne_Tableau_fils.VOI2, Ligne_Tableau_fils.[Ref Connecteur2], Ligne_Tableau_fils.[Ref Connecteur_Four2],  "
Sql = Sql & "Ligne_Tableau_fils.Long_Add2, Ligne_Tableau_fils.[Ref Clip2], Ligne_Tableau_fils.[Ref Clip Four2],  "
Sql = Sql & "Ligne_Tableau_fils.[Ref Joint2] , Ligne_Tableau_fils.[Ref Joint Four2],   "
Sql = Sql & " Ligne_Tableau_fils.PRECOG,  "
Sql = Sql & "Ligne_Tableau_fils.PRECO2, Ligne_Tableau_fils.PRECO, Ligne_Tableau_fils.Option, Ligne_Tableau_fils.Activer,  "
Sql = Sql & "Ligne_Tableau_fils.[Critères spécifiques],Ligne_Tableau_fils.COMMENTAIRES "

Sql = Sql & "FROM Ligne_Tableau_fils LEFT JOIN Ligne_Tableau_fils_Ecart "
Sql = Sql & "ON Ligne_Tableau_fils.Id = Ligne_Tableau_fils_Ecart.Id "
Sql = Sql & "WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & IdIndiceProjet & "  "
Sql = Sql & "AND Ligne_Tableau_fils_Ecart.id Is Null;"

Set RsAjouter = Con.OpenRecordSet(Sql)


Sql = "SELECT '' AS [Avant/Après], Ligne_Tableau_fils_Ecart.LIAI, Ligne_Tableau_fils_Ecart.DESIGNATION,  "
Sql = Sql & "Ligne_Tableau_fils_Ecart.FIL, Ligne_Tableau_fils_Ecart.SECT, Ligne_Tableau_fils_Ecart.TEINT,  "
Sql = Sql & "Ligne_Tableau_fils_Ecart.TEINT2, Ligne_Tableau_fils_Ecart.ISO, Ligne_Tableau_fils_Ecart.LONG,  "
Sql = Sql & "Ligne_Tableau_fils_Ecart.[LONG CP], Ligne_Tableau_fils_Ecart.COUPE, Ligne_Tableau_fils_Ecart.POS,  "
Sql = Sql & "Ligne_Tableau_fils_Ecart.[POS-OUT], Ligne_Tableau_fils_Ecart.FA, Ligne_Tableau_fils_Ecart.APP,  "
Sql = Sql & "Ligne_Tableau_fils_Ecart.VOI, Ligne_Tableau_fils_Ecart.[Ref Connecteur],  "
Sql = Sql & "Ligne_Tableau_fils_Ecart.[Ref Connecteur_Four], Ligne_Tableau_fils_Ecart.Long_Add,  "
Sql = Sql & "Ligne_Tableau_fils_Ecart.[Ref Clip], Ligne_Tableau_fils_Ecart.[Ref Clip Four],  "
Sql = Sql & "Ligne_Tableau_fils_Ecart.[Ref Joint], Ligne_Tableau_fils_Ecart.[Ref Joint four],  "
Sql = Sql & "Ligne_Tableau_fils_Ecart.POS2, Ligne_Tableau_fils_Ecart.[POS-OUT2], Ligne_Tableau_fils_Ecart.FA2,  "
Sql = Sql & "Ligne_Tableau_fils_Ecart.APP2, Ligne_Tableau_fils_Ecart.VOI2,  "
Sql = Sql & "Ligne_Tableau_fils_Ecart.[Ref Connecteur2], Ligne_Tableau_fils_Ecart.[Ref Connecteur_Four2],  "
Sql = Sql & "Ligne_Tableau_fils_Ecart.Long_Add2, Ligne_Tableau_fils_Ecart.[Ref Clip2], "
Sql = Sql & " Ligne_Tableau_fils_Ecart.[Ref Clip Four2], Ligne_Tableau_fils_Ecart.[Ref Joint2],  "
Sql = Sql & "Ligne_Tableau_fils_Ecart.[Ref Joint Four2], Ligne_Tableau_fils_Ecart.PRECO,  "
Sql = Sql & "Ligne_Tableau_fils_Ecart.PRECO2, Ligne_Tableau_fils_Ecart.PRECOG,  "
Sql = Sql & "Ligne_Tableau_fils_Ecart.OPTION, Ligne_Tableau_fils_Ecart.ACTIVER,  "
Sql = Sql & "Ligne_Tableau_fils_Ecart.[Critères spécifiques],  "
Sql = Sql & "Ligne_Tableau_fils_Ecart.Commentaires, '' AS Expr, Ligne_Tableau_fils.LIAI,  "
Sql = Sql & "Ligne_Tableau_fils.DESIGNATION, Ligne_Tableau_fils.FIL, Ligne_Tableau_fils.SECT,  "
Sql = Sql & "Ligne_Tableau_fils.TEINT, Ligne_Tableau_fils.TEINT2, Ligne_Tableau_fils.ISO,  "
Sql = Sql & "Ligne_Tableau_fils.LONG, Ligne_Tableau_fils.[LONG CP], Ligne_Tableau_fils.COUPE,  "
Sql = Sql & "Ligne_Tableau_fils.POS, Ligne_Tableau_fils.[POS-OUT], Ligne_Tableau_fils.FA,  "
Sql = Sql & "Ligne_Tableau_fils.APP, Ligne_Tableau_fils.VOI, Ligne_Tableau_fils.[Ref Connecteur],  "
Sql = Sql & "Ligne_Tableau_fils.[Ref Connecteur_Four], Ligne_Tableau_fils.Long_Add,  "
Sql = Sql & "Ligne_Tableau_fils.[Ref Clip], Ligne_Tableau_fils.[Ref Clip Four] ,  "
Sql = Sql & "Ligne_Tableau_fils.[Ref Joint], Ligne_Tableau_fils.[Ref Joint four], "
Sql = Sql & "Ligne_Tableau_fils.POS2, Ligne_Tableau_fils.[POS-OUT2], Ligne_Tableau_fils.FA2,  "
Sql = Sql & "Ligne_Tableau_fils.APP2, Ligne_Tableau_fils.VOI2, Ligne_Tableau_fils.[Ref Connecteur2],  "
Sql = Sql & "Ligne_Tableau_fils.[Ref Connecteur_Four2], Ligne_Tableau_fils.Long_Add2,  "
Sql = Sql & "Ligne_Tableau_fils.[Ref Clip2], Ligne_Tableau_fils.[Ref Clip Four2],  "
Sql = Sql & "Ligne_Tableau_fils.[Ref Joint2], Ligne_Tableau_fils.[Ref Joint Four2],  "
Sql = Sql & "Ligne_Tableau_fils.PRECO, Ligne_Tableau_fils.PRECO2, Ligne_Tableau_fils.PRECOG,  "
Sql = Sql & "Ligne_Tableau_fils.OPTION, Ligne_Tableau_fils.ACTIVER, Ligne_Tableau_fils.[Critères spécifiques],  "
Sql = Sql & "Ligne_Tableau_fils.Commentaires "

Sql = Sql & "FROM Ligne_Tableau_fils_Ecart INNER JOIN Ligne_Tableau_fils  "
Sql = Sql & "ON Ligne_Tableau_fils_Ecart.Id = Ligne_Tableau_fils.Id "
Sql = Sql & "WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & IdIndiceProjet & ";  "


Set RsModifier = Con.OpenRecordSet(Sql)

L = 0: C = 0
If MajEcartExcel(IdIndiceProjet, MyWorkbook, L, C, RsSuprimer, RsAjouter, RsModifier, "Tableau_fils_Ecart") = True Then boolSave = True
Sql = "SELECT Connecteurs_Ecart.CONNECTEUR, Connecteurs_Ecart.[O/N], Connecteurs_Ecart.DESIGNATION,  "
Sql = Sql & "Connecteurs_Ecart.CODE_APP, Connecteurs_Ecart.N°, Connecteurs_Ecart.POS,  "
Sql = Sql & "Connecteurs_Ecart.[POS-OUT], Connecteurs_Ecart.PRECO1, Connecteurs_Ecart.PRECO2,  "
Sql = Sql & "Connecteurs_Ecart.[100%], Connecteurs_Ecart.OPTION, Connecteurs_Ecart.Pylone,  "
Sql = Sql & "Connecteurs_Ecart.Colonne, Connecteurs_Ecart.Ligne, Connecteurs_Ecart.ACTIVER,  "
Sql = Sql & "Connecteurs_Ecart.RefBouchon, Connecteurs_Ecart.RefBouchonFour, Connecteurs_Ecart.ReFCapot,  "
Sql = Sql & "Connecteurs_Ecart.ReFCapotFour, Connecteurs_Ecart.RefVerrou, Connecteurs_Ecart.RefVerrouFour,  "
Sql = Sql & "Connecteurs_Ecart.RefConnecteurFour, Connecteurs_Ecart.LongueurF_Choix,Connecteurs_Ecart.COMMENTAIRES "
Sql = Sql & "FROM Connecteurs_Ecart LEFT JOIN Connecteurs ON Connecteurs_Ecart.Id = Connecteurs.Id "
Sql = Sql & "WHERE Connecteurs_Ecart.Id_IndiceProjet=" & IdIndiceProjet & "  "
Sql = Sql & "AND Connecteurs.id Is Null;"
Set RsSuprimer = Con.OpenRecordSet(Sql)

Sql = "SELECT Connecteurs.CONNECTEUR, Connecteurs.[O/N], Connecteurs.DESIGNATION, Connecteurs.CODE_APP,  "
Sql = Sql & "Connecteurs.N°, Connecteurs.POS, Connecteurs.[POS-OUT], Connecteurs.PRECO1,  "
Sql = Sql & "Connecteurs.PRECO2, Connecteurs.[100%], Connecteurs.OPTION, Connecteurs.Pylone,  "
Sql = Sql & "Connecteurs.Colonne, Connecteurs.Ligne, Connecteurs.ACTIVER, Connecteurs.RefBouchon,  "
Sql = Sql & "Connecteurs.RefBouchonFour, Connecteurs.ReFCapot, Connecteurs.ReFCapotFour,  "
Sql = Sql & "Connecteurs.RefVerrou, Connecteurs.RefVerrouFour, Connecteurs.RefConnecteurFour,  "
Sql = Sql & "Connecteurs.LongueurF_Choix,Connecteurs.COMMENTAIRES "

Sql = Sql & "FROM Connecteurs LEFT JOIN Connecteurs_Ecart ON Connecteurs.Id = Connecteurs_Ecart.Id "
Sql = Sql & "WHERE Connecteurs.Id_IndiceProjet=" & IdIndiceProjet & "  "
Sql = Sql & "AND Connecteurs_Ecart.id Is Null;"

Set RsAjouter = Con.OpenRecordSet(Sql)

Sql = "SELECT '' AS [Avant/Après], Connecteurs_Ecart.CONNECTEUR, Connecteurs_Ecart.[O/N], "
Sql = Sql & "Connecteurs_Ecart.DESIGNATION, Connecteurs_Ecart.CODE_APP, Connecteurs_Ecart.N°,  "
Sql = Sql & "Connecteurs_Ecart.POS, Connecteurs_Ecart.[POS-OUT], Connecteurs_Ecart.PRECO1,  "
Sql = Sql & "Connecteurs_Ecart.PRECO2, Connecteurs_Ecart.[100%], Connecteurs_Ecart.OPTION,  "
Sql = Sql & "Connecteurs_Ecart.Pylone, Connecteurs_Ecart.Colonne, Connecteurs_Ecart.Ligne,  "
Sql = Sql & "Connecteurs_Ecart.ACTIVER, Connecteurs_Ecart.RefBouchon, Connecteurs_Ecart.RefBouchonFour,  "
Sql = Sql & "Connecteurs_Ecart.ReFCapot, Connecteurs_Ecart.ReFCapotFour, Connecteurs_Ecart.RefVerrou,  "
Sql = Sql & "Connecteurs_Ecart.RefVerrouFour, Connecteurs_Ecart.RefConnecteurFour,  "
Sql = Sql & "Connecteurs_Ecart.LongueurF_Choix,Connecteurs_Ecart.COMMENTAIRES , '' AS Expr1, Connecteurs.CONNECTEUR, Connecteurs.[O/N],  "
Sql = Sql & "Connecteurs.DESIGNATION, Connecteurs.CODE_APP, Connecteurs.N°, Connecteurs.POS,  "
Sql = Sql & "Connecteurs.[POS-OUT], Connecteurs.PRECO1, Connecteurs.PRECO2, Connecteurs.[100%],  "
Sql = Sql & "Connecteurs.OPTION, Connecteurs.Pylone, Connecteurs.Colonne, Connecteurs.Ligne,  "
Sql = Sql & "Connecteurs.ACTIVER, Connecteurs.RefBouchon,  "
Sql = Sql & "Connecteurs.RefBouchonFour, Connecteurs.ReFCapot, Connecteurs.ReFCapotFour,  "
Sql = Sql & "Connecteurs.RefVerrou, Connecteurs.RefVerrouFour, Connecteurs.RefConnecteurFour,  "
Sql = Sql & "Connecteurs.LongueurF_Choix,Connecteurs.COMMENTAIRES "

Sql = Sql & "FROM Connecteurs INNER JOIN Connecteurs_Ecart ON Connecteurs.Id = Connecteurs_Ecart.Id "
Sql = Sql & "WHERE Connecteurs.Id_IndiceProjet=" & IdIndiceProjet & "  "



Set RsModifier = Con.OpenRecordSet(Sql)

L = 0: C = 0

If MajEcartExcel(IdIndiceProjet, MyWorkbook, L, C, RsSuprimer, RsAjouter, RsModifier, "Connecteurs_Ecart") = True Then boolSave = True

Sql = "SELECT T_Critères_Ecart.ACTIVER,T_Critères_Ecart.CODE_CRITERE, T_Critères_Ecart.CRITERES,T_Critères_Ecart.COMMENTAIRES "
Sql = Sql & "FROM T_Critères_Ecart LEFT JOIN T_Critères ON T_Critères_Ecart.Id = T_Critères.Id "
Sql = Sql & "WHERE T_Critères_Ecart.Id_IndiceProjet=" & IdIndiceProjet & "  "
Sql = Sql & "AND T_Critères.id Is Null;"


Set RsSuprimer = Con.OpenRecordSet(Sql)

Sql = "SELECT T_Critères.ACTIVER,T_Critères.CODE_CRITERE, T_Critères.CRITERES,T_Critères.COMMENTAIRES "
Sql = Sql & "FROM T_Critères LEFT JOIN T_Critères_Ecart ON T_Critères.Id =T_Critères_Ecart.Id "
Sql = Sql & "WHERE T_Critères.Id_IndiceProjet=" & IdIndiceProjet & "  "
Sql = Sql & "AND T_Critères_Ecart.id Is Null;"


Set RsAjouter = Con.OpenRecordSet(Sql)

Sql = "SELECT '' AS [Avant/Après], T_Critères_Ecart.ACTIVER,T_Critères_Ecart.CODE_CRITERE,  "
Sql = Sql & "T_Critères_Ecart.CRITERES,T_Critères_Ecart.COMMENTAIRES , '' AS Expr1, T_Critères.ACTIVER,T_Critères.CODE_CRITERE, T_Critères.CRITERES,T_Critères.COMMENTAIRES "
Sql = Sql & "FROM T_Critères INNER JOIN T_Critères_Ecart ON T_Critères.Id = T_Critères_Ecart.Id "
Sql = Sql & "WHERE T_Critères.Id_IndiceProjet=" & IdIndiceProjet & "  "
Sql = Sql & "ORDER BY T_Critères.CODE_CRITERE;"


Set RsModifier = Con.OpenRecordSet(Sql)

L = 0: C = 0
If MajEcartExcel(IdIndiceProjet, MyWorkbook, L, C, RsSuprimer, RsAjouter, RsModifier, "Critères_Ecart") = True Then boolSave = True
If boolSave = True Then
Sql = "SELECT RqCartouche.* "
Sql = Sql & "FROM RqCartouche "
 Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"
Set RsModifier = Con.OpenRecordSet(Sql)
If RsModifier.EOF = False Then
        PathPl = PathArchive(TableauPath("PathArchiveAutocad"), "" & RsModifier!Client, "" & RsModifier!CleAc, "" & RsModifier!Pieces, "Li", RsModifier.Fields("Li"), IdIndiceProjet, RsModifier.Fields("PI_Indice"), RsModifier.Fields("LI_Indice"), RsModifier!Version, True)
        PathPl = Replace(PathPl, RsModifier.Fields("Li") & "_" & Trim(RsModifier.Fields("LI_Indice")), "Ecart")
        Dim Fso As New FileSystemObject
        If Fso.FolderExists(PathPl) = False Then Fso.CreateFolder PathPl
        PathPl = PathPl & "\" & RsModifier.Fields("Li") & "_" & Trim(RsModifier.Fields("LI_Indice"))
RepriseSave:
        MyFormatDate = Format(Now, "yyyy-mm-dd-h-m-s")
        If Fso.FileExists(PathPl & "_Ecart_" & MyFormatDate & ".XLS") = True Then
        DoEvents
          GoTo RepriseSave
        End If
        MyWorkbook.SaveAs PathPl & "_Ecart_" & MyFormatDate, ReadOnlyRecommended:=True
        
        If IdFils <> 0 Then
        Sql = "SELECT RqCartouche.* "
        Sql = Sql & "FROM RqCartouche "
        Sql = Sql & "WHERE T_indiceProjet.Id=" & IdFils & ";"
        Set RsModifier = Con.OpenRecordSet(Sql)
         PathPl2 = PathArchive(PathArchiveAutocad, "" & RsModifier!Client, "" & RsModifier!CleAc, "" & RsModifier!Pieces, "Li", RsModifier.Fields("Li"), IdFils, RsModifier.Fields("PI_Indice"), RsModifier.Fields("LI_Indice"), RsModifier!Version, True)
         
         PathPl2 = Replace(PathPl2, RsModifier.Fields("Li") & "_" & Trim(RsModifier.Fields("LI_Indice")), "Ecart")
         If Fso.FolderExists(PathPl2) = False Then Fso.CreateFolder PathPl2
        PathPl2 = PathPl2 & "\" & RsModifier.Fields("Li") & "_" & Trim(RsModifier.Fields("LI_Indice"))
        PathPl2 = PathPl2 & "_Ecart_" & MyFormatDate
       Racourci "" & PathPl2, "" & PathPl & "_Ecart_" & MyFormatDate, "XLS"
    End If
End If
End If
MyWorkbook.Close False
MyExcel.Quit
Set MyExcel = Nothing
End Sub
Function sqlRange(MyRange, IndexLigne As Long, FROM, Optional txtStop As String)
Dim Sql1 As String
Dim Sql2 As String
Dim Sql1Val As String
Dim Sql2Val As String
Dim Sql3Val As String
Dim TabDouble
Dim I As Long
Sql1 = "INSERT INTO " & FROM & " (Job,"
Sql1Val = ""
Sql2Val = NmJob & ","
Sql3Val = ""
'Myrange.Application.Visible = True
For I = 1 To MyRange.Columns.Count
'     If FROM = "Xls_Composants" And I > 9 Then Exit For
    If Trim("" & MyRange(1, I)) = "" Then Exit For
DoEvents
    Sql1Val = Sql1Val & "[" & MyRange(1, I) & "],"
'   Myrange.Application.Visible = True
    If Trim("" & MyRange(IndexLigne, I)) = "" Then
        If MyRange(1, I) = "O/N" Then
             Sql2Val = Sql2Val & "0,"
        Else
            Sql2Val = Sql2Val & "NULL,"
        End If
    Else
        If MyRange(1, I) = "O/N" Or MyRange(1, I) = "ACTIVER" Then
'            If Left(UCase(Trim(Myrange(IndexLigne, I))), 1) = "N" Then Myrange(IndexLigne, I) = 0
'            If Myrange(IndexLigne, I) <> 0 Then Myrange(IndexLigne, I) = -1
''            If Left(UCase(MyRange(IndexLigne, I)), 1) <> "F" Then MyRange(IndexLigne, I) = 1
            If Abs(MyRange(IndexLigne, I)) <> 1 Then MyRange(IndexLigne, I) = 0
            
            Sql2Val = Sql2Val & "" & CInt(Trim(MyRange(IndexLigne, I))) & ","
        Else
            isnu = "0" & Replace(Trim(MyRange(IndexLigne, I)), ",", ".")
            TabDouble = Split(isnu & ".", ".")
            If IsNumeric(isnu) = True And (IsNumeric(TabDouble(1)) And TabDouble(1) <> "") Then
            MyRange(IndexLigne, I) = "'" & Replace(Trim("" & MyRange(IndexLigne, I)), ",", ".")
            End If
'            Myrange.Application.Visible = True
            Sql2Val = Sql2Val & "'" & MyReplace(Trim(MyRange(IndexLigne, I))) & "',"
        End If
    End If
   If Trim("" & txtStop) <> "" Then
    If UCase(Trim("" & MyRange(1, I))) = UCase(txtStop) Then
        Exit For
    End If
   
   End If
Next I
'If FROM = "Xls_Composants" Then
'Sql1Val = Sql1Val & "[Path],"
'Sql3Val = "NULL,"
'    For I = I To Myrange.Columns.Count
'        If Val((Trim("" & Myrange(IndexLigne, I).Value))) = 1 Then
'            Sql3Val = "'" & MyReplace(Trim(Myrange(1, I))) & "',"
'            Exit For
'        End If
'    Next I
'End If

Sql1Val = Left(Sql1Val, Len(Sql1Val) - 1)
Sql2Val = Sql2Val & Sql3Val
Sql2Val = Left(Sql2Val, Len(Sql2Val) - 1)
sqlRange = Sql1 & Sql1Val & ") Values(" & Sql2Val & ");"
End Function

Sub CeerFichierXls(Xls As String)
Dim Sql As String
Dim Rs As Recordset
Dim IdProjet As Long
Dim IdIndice As Long
DoEvents

Dim MyExcel As New EXCEL.Application
MyExcel.DisplayAlerts = False
Dim MyClasseur As EXCEL.Workbook
Dim MySheet As EXCEL.Worksheet
Dim MyRange As EXCEL.Range
Set MyClasseur = MyExcel.Workbooks.Open(TableauPath.Item("PathModelXls") & "Ligne_Tableau_fils.xlt")
MyExcel.DisplayAlerts = False
MyClasseur.SaveAs Xls, ReadOnlyRecommended:=True
MyExcel.DisplayAlerts = True
'MyExcel.Visible = True
End Sub
Function MajEcartConnecteur(IdIndiceProjet As Long, MyWorkbook As EXCEL.Workbook, L As Long, C As Long, RsSuprimer As Recordset, RsAjouter As Recordset, RsModifier As Recordset) As Boolean
Dim Sql As String
Dim Rs As Recordset

Dim MySheet As EXCEL.Worksheet
Dim MyRange As Range
Sql = "SELECT Connecteurs_Ecart.* "
Sql = Sql & "FROM Connecteurs_Ecart LEFT JOIN Connecteurs  "
Sql = Sql & "ON (Connecteurs_Ecart.Id_IndiceProjet = Connecteurs.Id_IndiceProjet)  "
Sql = Sql & "AND (Connecteurs_Ecart.N° = Connecteurs.N°) "
Sql = Sql & "WHERE Connecteurs_Ecart.Id_IndiceProjet=" & IdIndiceProjet & "  "
Sql = Sql & "AND Connecteurs.CODE_APP Is Null;"
 
 

Set Rs = Con.OpenRecordSet(Sql)
Set MySheet = IsertSheet(MyWorkbook, "Connecteur_Ecart")
Set MyRange = MySheet.Cells(1, 1).CurrentRegion
L = MyRange.Rows.Count
If L > 1 Then L = L + 1
If Rs.EOF = False Then
MajEcartConnecteur = True
MyRange(MyRange.Rows.Count, 1) = "Enregistrement Suprimer"
FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L, Rs.Fields.Count).Address), 40, True, True, xlCenter, xlCenter


Set MyRange = MySheet.Cells(1, 1).CurrentRegion
L = MyRange.Rows.Count + 1
For I = 0 To Rs.Fields.Count - 1
    MyRange(L, I + 1) = Rs(I).Name
Next
FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L, Rs.Fields.Count).Address), 15, False, True, xlCenter, xlCenter

While Rs.EOF = False
L = L + 1
   For I = 0 To Rs.Fields.Count - 1
    MyRange(L, I + 1) = "" & Rs(I).Value
    Next
    FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L, Rs.Fields.Count).Address), 2, False, True, xlCenter, xlCenter

    Rs.MoveNext
Wend
End If
Set Rs = Con.CloseRecordSet(Rs)
Sql = "SELECT Connecteurs.* "
Sql = Sql & "FROM Connecteurs LEFT JOIN Connecteurs_Ecart  "
Sql = Sql & "ON (Connecteurs.N° = Connecteurs_Ecart.N°)  "
Sql = Sql & "AND (Connecteurs.Id_IndiceProjet = Connecteurs_Ecart.Id_IndiceProjet) "
Sql = Sql & "WHERE Connecteurs.Id_IndiceProjet=" & IdIndiceProjet & "  "
Sql = Sql & "AND Connecteurs_Ecart.CODE_APP Is Null;"


Set Rs = Con.OpenRecordSet(Sql)
Set MySheet = IsertSheet(MyWorkbook, "Connecteur_Ecart")
Set MyRange = MySheet.Cells(1, 1).CurrentRegion
L = MyRange.Rows.Count
If L > 1 Then L = L + 1
If Rs.EOF = False Then
MajEcartConnecteur = True
MyRange(L, 1) = "Enregistrement Ajouter"
FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L, Rs.Fields.Count).Address), 40, True, True, xlCenter, xlCenter


Set MyRange = MySheet.Cells(1, 1).CurrentRegion
L = MyRange.Rows.Count + 1

For I = 0 To Rs.Fields.Count - 1
    MyRange(L, I + 1) = Rs(I).Name
Next
FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L, Rs.Fields.Count).Address), 15, False, True, xlCenter, xlCenter

While Rs.EOF = False
L = L + 1
   For I = 0 To Rs.Fields.Count - 1
    MyRange(L, I + 1) = "" & Rs(I).Value
    Next
    FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L, Rs.Fields.Count).Address), 2, False, True, xlCenter, xlCenter

    Rs.MoveNext
Wend



End If
Set Rs = Con.CloseRecordSet(Rs)
Sql = "SELECT Connecteurs_Ecart.*, Connecteurs.* "
Sql = Sql & "FROM Connecteurs INNER JOIN Connecteurs_Ecart ON (Connecteurs.N° = Connecteurs_Ecart.N°) AND (Connecteurs.Id_IndiceProjet = Connecteurs_Ecart.Id_IndiceProjet) "
Sql = Sql & "WHERE Connecteurs.Id_IndiceProjet=" & IdIndiceProjet & "  "


Set Rs = Con.OpenRecordSet(Sql)
Set MySheet = IsertSheet(MyWorkbook, "Connecteur_Ecart")
Set MyRange = MySheet.Cells(1, 1).CurrentRegion
L = MyRange.Rows.Count
If L > 1 Then L = L + 1
If Rs.EOF = False Then
MyRange(L, 1) = "Enregistrement Modifier"
FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L, (Rs.Fields.Count / 2)).Address), 40, True, True, xlCenter, xlCenter


Set MyRange = MySheet.Cells(1, 1).CurrentRegion
L = MyRange.Rows.Count + 1
For I = 0 To (Rs.Fields.Count / 2) - 1
    MyRange(L, I + 1) = Rs(I).Name
Next
FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L, Rs.Fields.Count / 2).Address), 15, False, True, xlCenter, xlCenter

While Rs.EOF = False
L = L + 1
modifire = False
   For I = 0 To (Rs.Fields.Count - 2) / 2
    If "" & Rs(I).Value <> "" & Rs(13 + I).Value Then
        If I > 0 Then
        MajEcartConnecteur = True
        modifire = True
        Exit For
        End If
    End If
   
    Next
    If modifire = True Then
   For I = 0 To (Rs.Fields.Count - 2) / 2
    MyRange(L, I + 1) = "" & Rs(I).Value & Chr(10) & "" & Rs(13 + I).Value
    If "" & Rs(I).Value <> "" & Rs(13 + I).Value And I > 0 Then
        MyRange(L, I + 1).Font.ColorIndex = 3
       
    End If
    
    Next
          FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L, Rs.Fields.Count / 2).Address), 2, False, True, xlCenter, xlCenter
    Else
        L = L - 1
    End If
  

    Rs.MoveNext
Wend
End If

Set Rs = Con.CloseRecordSet(Rs)
Set MyRange = MyRange(1, 1).CurrentRegion
If MyRange.Rows.Count = 1 Then Exit Function

Sql = "SELECT  T_indiceProjet.Ensemble, T_indiceProjet.Equipement,[Li] & '_' & [LI_Indice] AS Liste, "
Sql = Sql & "[PI] & '_' & [PI_Indice] AS Piece, T_indiceProjet.Client, T_indiceProjet.CleAc "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"
Dim RsEntetePage As Recordset
Set RsEntetePage = Con.OpenRecordSet(Sql)
If PortraitPaysage = 0 Then PortraitPaysage = 2
MiseEnPage MySheet, MyRange, "Affaire : " & RsEntetePage!CleAc & Chr(10) & _
    "Câblage : " & Replace("" & RsEntetePage!Ensemble, vbCrLf, " ") & Chr(10) & _
     "Pièce : " & RsEntetePage!Piece, vbCrLf & "Equipement : " & Replace("" & RsEntetePage!Equipement, vbCrLf, " ") & Chr(10) & _
     "Liste : " & RsEntetePage!Liste _
    , "ENCELADE" & Chr(10) & "Client : " & RsEntetePage!Client & Chr(10) & Format(Date, "dd-mmm-yyyy"), "", "&P/&N", "", 80, "", False, PortraitPaysage, False, True, True

    MaJEncadreXls MyRange, xlThin, xlThin, xlHairline, xlHairline


Set RsEntetePage = Con.CloseRecordSet(RsEntetePage)
End Function
Function MajEcartCritaire(IdIndiceProjet As Long, MyWorkbook As EXCEL.Workbook, L As Long, C As Long, RsSuprimer As Recordset, RsAjouter As Recordset, RsModifier As Recordset) As Boolean
Dim Sql As String
Dim Rs As Recordset

Dim MySheet As EXCEL.Worksheet
Dim MyRange As Range
Sql = "SELECT T_Critères_Ecart.*  "
Sql = Sql & "FROM T_Critères_Ecart LEFT JOIN T_Critères ON (T_Critères_Ecart.Id_IndiceProjet = T_Critères.Id_IndiceProjet) "
Sql = Sql & "AND (T_Critères_Ecart.CODE_CRITERE = T_Critères.CODE_CRITERE) "
Sql = Sql & "WHERE T_Critères_Ecart.Id_IndiceProjet=" & IdIndiceProjet & "  "
Sql = Sql & "AND T_Critères.CODE_CRITERE Is Null;"

Set Rs = Con.OpenRecordSet(Sql)
Set MySheet = IsertSheet(MyWorkbook, "Critères_Ecart")
Set MyRange = MySheet.Cells(1, 1).CurrentRegion
L = MyRange.Rows.Count
If L > 1 Then L = L + 1
If Rs.EOF = False Then
MajEcartCritaire = True
MyRange(MyRange.Rows.Count, 1) = "Enregistrement Suprimer"
FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L, Rs.Fields.Count).Address), 40, True, True, xlCenter, xlCenter


Set MyRange = MySheet.Cells(1, 1).CurrentRegion
L = MyRange.Rows.Count + 1
For I = 0 To Rs.Fields.Count - 1
    MyRange(L, I + 1) = Rs(I).Name
Next
FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L, Rs.Fields.Count).Address), 15, False, True, xlCenter, xlCenter

While Rs.EOF = False
L = L + 1
   For I = 0 To Rs.Fields.Count - 1
    MyRange(L, I + 1) = "" & Rs(I).Value
    Next
    FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L, Rs.Fields.Count).Address), 2, False, True, xlCenter, xlCenter

    Rs.MoveNext
Wend
End If
Set Rs = Con.CloseRecordSet(Rs)
Sql = "SELECT T_Critères.* "
Sql = Sql & "FROM T_Critères LEFT JOIN T_Critères_Ecart ON (T_Critères.Id_IndiceProjet =  "
Sql = Sql & "T_Critères_Ecart.Id_IndiceProjet) AND (T_Critères.CODE_CRITERE = T_Critères_Ecart.CODE_CRITERE)"
Sql = Sql & "WHERE T_Critères.Id_IndiceProjet=" & IdIndiceProjet & "  "
Sql = Sql & "AND T_Critères_Ecart.CODE_CRITERE Is Null;"

Set Rs = Con.OpenRecordSet(Sql)
Set MySheet = IsertSheet(MyWorkbook, "Critères_Ecart")
Set MyRange = MySheet.Cells(1, 1).CurrentRegion
L = MyRange.Rows.Count
If L > 1 Then L = L + 1
If Rs.EOF = False Then
MajEcartCritaire = True
MyRange(L, 1) = "Enregistrement Ajouter"
FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L, Rs.Fields.Count).Address), 40, True, True, xlCenter, xlCenter


Set MyRange = MySheet.Cells(1, 1).CurrentRegion
L = MyRange.Rows.Count + 1

For I = 0 To Rs.Fields.Count - 1
    MyRange(L, I + 1) = Rs(I).Name
Next
FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L, Rs.Fields.Count).Address), 15, False, True, xlCenter, xlCenter

While Rs.EOF = False
L = L + 1
   For I = 0 To Rs.Fields.Count - 1
    MyRange(L, I + 1) = "" & Rs(I).Value
    Next
    FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L, Rs.Fields.Count).Address), 2, False, True, xlCenter, xlCenter

    Rs.MoveNext
Wend



End If
Set Rs = Con.CloseRecordSet(Rs)
Sql = "SELECT  T_Critères_Ecart.*,T_Critères.* "
Sql = Sql & "FROM T_Critères INNER JOIN T_Critères_Ecart ON (T_Critères.Id_IndiceProjet =  "
Sql = Sql & "T_Critères_Ecart.Id_IndiceProjet) AND (T_Critères.CODE_CRITERE = T_Critères_Ecart.CODE_CRITERE) "
Sql = Sql & "WHERE T_Critères.Id_IndiceProjet=" & IdIndiceProjet & "  "
Sql = Sql & "ORDER BY T_Critères.CODE_CRITERE;"

Set Rs = Con.OpenRecordSet(Sql)
Set MySheet = IsertSheet(MyWorkbook, "Critères_Ecart")
Set MyRange = MySheet.Cells(1, 1).CurrentRegion
L = MyRange.Rows.Count
If L > 1 Then L = L + 1
If Rs.EOF = False Then
MyRange(L, 1) = "Enregistrement Modifier"
FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L, (Rs.Fields.Count / 2)).Address), 40, True, True, xlCenter, xlCenter


Set MyRange = MySheet.Cells(1, 1).CurrentRegion
L = MyRange.Rows.Count + 1
For I = 0 To (Rs.Fields.Count / 2) - 1
    MyRange(L, I + 1) = Rs(I).Name
Next
FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L, Rs.Fields.Count / 2).Address), 15, False, True, xlCenter, xlCenter

While Rs.EOF = False
L = L + 1
modifire = False
   For I = 0 To (Rs.Fields.Count - 2) / 2
    If "" & Rs(I).Value <> "" & Rs(4 + I).Value Then
        If I > 0 Then
        modifire = True
        MajEcartCritaire = True
        Exit For
        End If
    End If
   
    Next
    If modifire = True Then
   For I = 0 To (Rs.Fields.Count - 2) / 2
    MyRange(L, I + 1) = "" & Rs(I).Value & Chr(10) & "" & Rs(4 + I).Value
    If "" & Rs(I).Value <> "" & Rs(4 + I).Value And I > 0 Then
        MyRange(L, I + 1).Font.ColorIndex = 3
       
    End If
    
    Next
          FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L, Rs.Fields.Count / 2).Address), 2, False, True, xlCenter, xlCenter
    Else
        L = L - 1
    End If
  

    Rs.MoveNext
Wend
End If
Set Rs = Con.CloseRecordSet(Rs)
Set MyRange = MyRange(1, 1).CurrentRegion
If MyRange.Rows.Count = 1 Then Exit Function
Sql = "SELECT  T_indiceProjet.Ensemble, T_indiceProjet.Equipement,[Li] & '_' & [LI_Indice] AS Liste, "
Sql = Sql & "[PI] & '_' & [PI_Indice] AS Piece, T_indiceProjet.Client, T_indiceProjet.CleAc "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"
Dim RsEntetePage As Recordset
Set RsEntetePage = Con.OpenRecordSet(Sql)
Set MyRange = MyRange(1, 1).CurrentRegion
If PortraitPaysage = 0 Then PortraitPaysage = 2
MiseEnPage MySheet, MyRange, "Affaire : " & RsEntetePage!CleAc & Chr(10) & _
    "Câblage : " & Replace("" & RsEntetePage!Ensemble, vbCrLf, " ") & Chr(10) & _
     "Pièce : " & RsEntetePage!Piece, vbCrLf & "Equipement : " & Replace("" & RsEntetePage!Equipement, vbCrLf, " ") & Chr(10) & _
     "Liste : " & RsEntetePage!Liste _
    , "ENCELADE" & Chr(10) & "Client : " & RsEntetePage!Client & Chr(10) & Format(Date, "dd-mmm-yyyy"), "", "&P/&N", "", 80, "", False, PortraitPaysage, False, True, True
    
     MaJEncadreXls MyRange, xlThin, xlThin, xlHairline, xlHairline

Set RsEntetePage = Con.CloseRecordSet(RsEntetePage)
End Function

Function MajEcartFils(IdIndiceProjet As Long, MyWorkbook As EXCEL.Workbook, L As Long, C As Long, RsSuprimer As Recordset, RsAjouter As Recordset, RsModifier As Recordset) As Boolean
Dim Sql As String
Dim Rs As Recordset

Dim MySheet As EXCEL.Worksheet
Dim MyRange As Range
Sql = "SELECT Ligne_Tableau_fils_Ecart.*  "
Sql = Sql & "FROM Ligne_Tableau_fils_Ecart LEFT JOIN Ligne_Tableau_fils   "
Sql = Sql & "ON (Ligne_Tableau_fils_Ecart.FIL = Ligne_Tableau_fils.FIL)   "
Sql = Sql & "AND (Ligne_Tableau_fils_Ecart.Id_IndiceProjet = Ligne_Tableau_fils.Id_IndiceProjet)  "
Sql = Sql & "WHERE Ligne_Tableau_fils.FIL Is Null AND Ligne_Tableau_fils_Ecart.Id_IndiceProjet=" & IdIndiceProjet & ";  "

 
 

Set Rs = Con.OpenRecordSet(Sql)
Set MySheet = IsertSheet(MyWorkbook, "Tableau_Fils_Ecart")
Set MyRange = MySheet.Cells(1, 1).CurrentRegion
L = MyRange.Rows.Count
If L > 1 Then L = L + 1
If Rs.EOF = False Then
MajEcartFils = True
MyRange(MyRange.Rows.Count, 1) = "Enregistrement Suprimer"
FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L, Rs.Fields.Count).Address), 40, True, True, xlCenter, xlCenter


Set MyRange = MySheet.Cells(1, 1).CurrentRegion
L = MyRange.Rows.Count + 1
For I = 0 To Rs.Fields.Count - 1
    MyRange(L, I + 1) = Rs(I).Name
Next
FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L, Rs.Fields.Count).Address), 15, False, True, xlCenter, xlCenter

While Rs.EOF = False
L = L + 1
   For I = 0 To Rs.Fields.Count - 1
    MyRange(L, I + 1) = "" & Rs(I).Value
    Next
    FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L, Rs.Fields.Count).Address), 2, False, True, xlCenter, xlCenter

    Rs.MoveNext
Wend
End If
Set Rs = Con.CloseRecordSet(Rs)
Sql = "SELECT Ligne_Tableau_fils.* "
Sql = Sql & "FROM Ligne_Tableau_fils LEFT JOIN Ligne_Tableau_fils_Ecart  "
Sql = Sql & "ON (Ligne_Tableau_fils.FIL = Ligne_Tableau_fils_Ecart.FIL)  "
Sql = Sql & "AND (Ligne_Tableau_fils.Id_IndiceProjet = Ligne_Tableau_fils_Ecart.Id_IndiceProjet) "
Sql = Sql & "WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & IdIndiceProjet & "  "
Sql = Sql & "AND Ligne_Tableau_fils_Ecart.LIAI Is Null;"


Set Rs = Con.OpenRecordSet(Sql)
Set MySheet = IsertSheet(MyWorkbook, "Tableau_Fils_Ecart")
Set MyRange = MySheet.Cells(1, 1).CurrentRegion
L = MyRange.Rows.Count
If L > 1 Then L = L + 1
If Rs.EOF = False Then
MajEcartFils = True
MyRange(L, 1) = "Enregistrement Ajouter"
FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L, Rs.Fields.Count).Address), 40, True, True, xlCenter, xlCenter


Set MyRange = MySheet.Cells(1, 1).CurrentRegion
L = MyRange.Rows.Count + 1

For I = 0 To Rs.Fields.Count - 1
    MyRange(L, I + 1) = Rs(I).Name
Next
FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L, Rs.Fields.Count).Address), 15, False, True, xlCenter, xlCenter

While Rs.EOF = False
L = L + 1
   For I = 0 To Rs.Fields.Count - 1
    MyRange(L, I + 1) = "" & Rs(I).Value
    Next
    FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L, Rs.Fields.Count).Address), 2, False, True, xlCenter, xlCenter

    Rs.MoveNext
Wend



End If
Set Rs = Con.CloseRecordSet(Rs)
Sql = "SELECT Ligne_Tableau_fils_Ecart.*, Ligne_Tableau_fils.* "
Sql = Sql & "FROM Ligne_Tableau_fils INNER JOIN Ligne_Tableau_fils_Ecart  "
Sql = Sql & "ON (Ligne_Tableau_fils.Id_IndiceProjet = Ligne_Tableau_fils_Ecart.Id_IndiceProjet)  "
Sql = Sql & "AND (Ligne_Tableau_fils.FIL = Ligne_Tableau_fils_Ecart.FIL) "
Sql = Sql & "WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & IdIndiceProjet & ";  "
 


Set Rs = Con.OpenRecordSet(Sql)
Set MySheet = IsertSheet(MyWorkbook, "Tableau_Fils_Ecart")
Set MyRange = MySheet.Cells(1, 1).CurrentRegion
L = MyRange.Rows.Count
If L > 1 Then L = L + 1
If Rs.EOF = False Then
MyRange(L, 1) = "Enregistrement Modifier"
FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L, (Rs.Fields.Count / 2)).Address), 40, True, True, xlCenter, xlCenter


Set MyRange = MySheet.Cells(1, 1).CurrentRegion
L = MyRange.Rows.Count + 1
For I = 0 To (Rs.Fields.Count / 2) - 1
    MyRange(L, I + 1) = Rs(I).Name
Next
FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L, Rs.Fields.Count / 2).Address), 15, False, True, xlCenter, xlCenter

While Rs.EOF = False
L = L + 1
modifire = False
   For I = 0 To (Rs.Fields.Count - 2) / 2
    If "" & Rs(I).Value <> "" & Rs(24 + I).Value Then
        If I > 0 Then
        MajEcartFils = True
        modifire = True
        Exit For
        End If
    End If
   
    Next
    If modifire = True Then
   For I = 0 To (Rs.Fields.Count - 2) / 2
    MyRange(L, I + 1) = "" & Rs(I).Value & Chr(10) & "" & Rs(24 + I).Value
    If "" & Rs(I).Value <> "" & Rs(24 + I).Value And I > 0 Then
        MyRange(L, I + 1).Font.ColorIndex = 3
       
    End If
    
    Next
          FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L, Rs.Fields.Count / 2).Address), 2, False, True, xlCenter, xlCenter
    Else
        L = L - 1
    End If
  

    Rs.MoveNext
Wend
End If
Set Rs = Con.CloseRecordSet(Rs)
Set MyRange = MyRange(1, 1).CurrentRegion
If MyRange.Rows.Count = 1 Then Exit Function

Sql = "SELECT  T_indiceProjet.Ensemble, T_indiceProjet.Equipement,[Li] & '_' & [LI_Indice] AS Liste, "
Sql = Sql & "[PI] & '_' & [PI_Indice] AS Piece, T_indiceProjet.Client, T_indiceProjet.CleAc "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"
Dim RsEntetePage As Recordset
Set RsEntetePage = Con.OpenRecordSet(Sql)

Set MyRange = MyRange(1, 1).CurrentRegion
If PortraitPaysage = 0 Then PortraitPaysage = 2
MiseEnPage MySheet, MyRange, "Affaire : " & RsEntetePage!CleAc & Chr(10) & _
    "Câblage : " & Replace("" & RsEntetePage!Ensemble, vbCrLf, " ") & Chr(10) & _
     "Pièce : " & RsEntetePage!Piece, vbCrLf & "Equipement : " & Replace("" & RsEntetePage!Equipement, vbCrLf, " ") & Chr(10) & _
     "Liste : " & RsEntetePage!Liste _
    , "ENCELADE" & Chr(10) & "Client : " & RsEntetePage!Client & Chr(10) & Format(Date, "dd-mmm-yyyy"), "", "&P/&N", "", 80, "", False, PortraitPaysage, False, True, True

    MaJEncadreXls MyRange, xlThin, xlThin, xlHairline, xlHairline

Set RsEntetePage = Con.CloseRecordSet(RsEntetePage)
End Function
Function MajEcartNota(IdIndiceProjet As Long, MyWorkbook As EXCEL.Workbook, L As Long, C As Long, RsSuprimer As Recordset, RsAjouter As Recordset, RsModifier As Recordset) As Boolean
Dim Sql As String
Dim Rs As Recordset

Dim MySheet As EXCEL.Worksheet
Dim MyRange As Range
Sql = "SELECT Nota_Ecart.* "
Sql = Sql & "FROM Nota_Ecart LEFT JOIN Nota ON (Nota_Ecart.NUMNOTA = Nota.NUMNOTA) AND (Nota_Ecart.Id_IndiceProjet = Nota.Id_IndiceProjet) "
Sql = Sql & "WHERE Nota_Ecart.Id_IndiceProjet=" & IdIndiceProjet & " "
Sql = Sql & "AND Nota.NUMNOTA Is Null;"


Set Rs = Con.OpenRecordSet(Sql)
Set MySheet = IsertSheet(MyWorkbook, "Notas_Ecart")
Set MyRange = MySheet.Cells(1, 1).CurrentRegion
L = MyRange.Rows.Count
If L > 1 Then L = L + 1
If Rs.EOF = False Then
MajEcartNota = True
MyRange(MyRange.Rows.Count, 1) = "Enregistrement Suprimer"
FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L, Rs.Fields.Count).Address), 40, True, True, xlCenter, xlCenter


Set MyRange = MySheet.Cells(1, 1).CurrentRegion
L = MyRange.Rows.Count + 1
For I = 0 To Rs.Fields.Count - 1
    MyRange(L, I + 1) = Rs(I).Name
Next
FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L, Rs.Fields.Count).Address), 15, False, True, xlCenter, xlCenter

While Rs.EOF = False
L = L + 1
   For I = 0 To Rs.Fields.Count - 1
    MyRange(L, I + 1) = "" & Rs(I).Value
    Next
    FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L, Rs.Fields.Count).Address), 2, False, True, xlCenter, xlCenter

    Rs.MoveNext
Wend
End If
Set Rs = Con.CloseRecordSet(Rs)
Sql = "SELECT Nota.* "
Sql = Sql & "FROM Nota LEFT JOIN Nota_Ecart  "
Sql = Sql & "ON (Nota.NUMNOTA = Nota_Ecart.NUMNOTA)  "
Sql = Sql & "AND (Nota.Id_IndiceProjet = Nota_Ecart.Id_IndiceProjet) "
Sql = Sql & "WHERE Nota.Id_IndiceProjet=" & IdIndiceProjet & "  "
Sql = Sql & "AND Nota_Ecart.NUMNOTA Is Null;"

Set Rs = Con.OpenRecordSet(Sql)
Set MySheet = IsertSheet(MyWorkbook, "Notas_Ecart")
Set MyRange = MySheet.Cells(1, 1).CurrentRegion
L = MyRange.Rows.Count
If L > 1 Then L = L + 1
If Rs.EOF = False Then
MajEcartNota = True
MyRange(L, 1) = "Enregistrement Ajouter"
FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L, Rs.Fields.Count).Address), 40, True, True, xlCenter, xlCenter


Set MyRange = MySheet.Cells(1, 1).CurrentRegion
L = MyRange.Rows.Count + 1

For I = 0 To Rs.Fields.Count - 1
    MyRange(L, I + 1) = Rs(I).Name
Next
FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L, Rs.Fields.Count).Address), 15, False, True, xlCenter, xlCenter

While Rs.EOF = False
L = L + 1
   For I = 0 To Rs.Fields.Count - 1
    MyRange(L, I + 1) = "" & Rs(I).Value
    Next
    FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L, Rs.Fields.Count).Address), 2, False, True, xlCenter, xlCenter

    Rs.MoveNext
Wend



End If
Set Rs = Con.CloseRecordSet(Rs)
Sql = "SELECT Nota.*, Nota_Ecart.* "
Sql = Sql & "FROM Nota INNER JOIN Nota_Ecart  "
Sql = Sql & "ON (Nota.NUMNOTA = Nota_Ecart.NUMNOTA)  "
Sql = Sql & "AND (Nota.Id_IndiceProjet = Nota_Ecart.Id_IndiceProjet) "
Sql = Sql & "WHERE Nota_Ecart.Id_IndiceProjet=" & IdIndiceProjet & ";  "



Set Rs = Con.OpenRecordSet(Sql)
Set MySheet = IsertSheet(MyWorkbook, "Notas_Ecart")
Set MyRange = MySheet.Cells(1, 1).CurrentRegion
L = MyRange.Rows.Count
If L > 1 Then L = L + 1
If Rs.EOF = False Then
MyRange(L, 1) = "Enregistrement Modifier"
FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L, (Rs.Fields.Count / 2)).Address), 40, True, True, xlCenter, xlCenter


Set MyRange = MySheet.Cells(1, 1).CurrentRegion
L = MyRange.Rows.Count + 1
For I = 0 To (Rs.Fields.Count / 2) - 1
    MyRange(L, I + 1) = Rs(I).Name
Next
FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L, Rs.Fields.Count / 2).Address), 15, False, True, xlCenter, xlCenter

While Rs.EOF = False
L = L + 1
modifire = False
   For I = 0 To (Rs.Fields.Count - 2) / 2
    If "" & Rs(I).Value <> "" & Rs(4 + I).Value Then
        If I > 0 Then
        modifire = True
        MajEcartNota = True
        Exit For
        End If
    End If
   
    Next
    If modifire = True Then
   For I = 0 To (Rs.Fields.Count - 2) / 2
    MyRange(L, I + 1) = "" & Rs(I).Value & Chr(10) & "" & Rs(4 + I).Value
    If "" & Rs(I).Value <> "" & Rs(4 + I).Value And I > 0 Then
        MyRange(L, I + 1).Font.ColorIndex = 3
       
    End If
    
    Next
          FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L, Rs.Fields.Count / 2).Address), 2, False, True, xlCenter, xlCenter
    Else
        L = L - 1
    End If
  

    Rs.MoveNext
Wend
End If
Set Rs = Con.CloseRecordSet(Rs)
Set MyRange = MyRange(1, 1).CurrentRegion
If MyRange.Rows.Count = 1 Then Exit Function

Sql = "SELECT  T_indiceProjet.Ensemble, T_indiceProjet.Equipement,[Li] & '_' & [LI_Indice] AS Liste, "
Sql = Sql & "[PI] & '_' & [PI_Indice] AS Piece, T_indiceProjet.Client, T_indiceProjet.CleAc "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"
Dim RsEntetePage As Recordset
Set RsEntetePage = Con.OpenRecordSet(Sql)
Set MyRange = MyRange(1, 1).CurrentRegion
If PortraitPaysage = 0 Then PortraitPaysage = 2
MiseEnPage MySheet, MyRange, "Affaire : " & RsEntetePage!CleAc & Chr(10) & _
    "Câblage : " & Replace("" & RsEntetePage!Ensemble, vbCrLf, " ") & Chr(10) & _
     "Pièce : " & RsEntetePage!Piece, vbCrLf & "Equipement : " & Replace("" & RsEntetePage!Equipement, vbCrLf, " ") & Chr(10) & _
     "Liste : " & RsEntetePage!Liste _
    , "ENCELADE" & Chr(10) & "Client : " & RsEntetePage!Client & Chr(10) & Format(Date, "dd-mmm-yyyy"), "", "&P/&N", "", 80, "", False, PortraitPaysage, False, True, True

    MaJEncadreXls MyRange, xlThin, xlThin, xlHairline, xlHairline


Set RsEntetePage = Con.CloseRecordSet(RsEntetePage)
End Function
Function MajEcartComposants(IdIndiceProjet As Long, MyWorkbook As EXCEL.Workbook, L As Long, C As Long, RsSuprimer As Recordset, RsAjouter As Recordset, RsModifier As Recordset) As Boolean
Dim Sql As String
Dim Rs As Recordset

Dim MySheet As EXCEL.Worksheet
Dim MyRange As Range
Sql = "SELECT Composants_Ecart.* "
Sql = Sql & "FROM Composants_Ecart LEFT JOIN Composants  "
Sql = Sql & "ON (Composants_Ecart.NUMCOMP = Composants.NUMCOMP)  "
Sql = Sql & "AND (Composants_Ecart.Id_IndiceProjet = Composants.Id_IndiceProjet) "
Sql = Sql & "WHERE Composants.NUMCOMP Is Null "
Sql = Sql & "AND Composants_Ecart.Id_IndiceProjet=" & IdIndiceProjet & " ;"
 

Set Rs = Con.OpenRecordSet(Sql)
Set MySheet = IsertSheet(MyWorkbook, "Composants_Ecart")
Set MyRange = MySheet.Cells(1, 1).CurrentRegion
L = MyRange.Rows.Count
If L > 1 Then L = L + 1
If Rs.EOF = False Then
MajEcartComposants = True
MyRange(MyRange.Rows.Count, 1) = "Enregistrement Suprimer"
FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L, Rs.Fields.Count).Address), 40, True, True, xlCenter, xlCenter


Set MyRange = MySheet.Cells(1, 1).CurrentRegion
L = MyRange.Rows.Count + 1
For I = 0 To Rs.Fields.Count - 1
    MyRange(L, I + 1) = Rs(I).Name
Next
FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L, Rs.Fields.Count).Address), 15, False, True, xlCenter, xlCenter

While Rs.EOF = False
L = L + 1
   For I = 0 To Rs.Fields.Count - 1
    MyRange(L, I + 1) = "" & Rs(I).Value
    Next
    FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L, Rs.Fields.Count).Address), 2, False, True, xlCenter, xlCenter

    Rs.MoveNext
Wend
End If
Set Rs = Con.CloseRecordSet(Rs)
Sql = "SELECT Composants.* "
Sql = Sql & "FROM Composants LEFT JOIN Composants_Ecart  "
Sql = Sql & "ON (Composants.NUMCOMP = Composants_Ecart.NUMCOMP)  "
Sql = Sql & "AND (Composants.Id_IndiceProjet = Composants_Ecart.Id_IndiceProjet) "
Sql = Sql & "WHERE Composants.Id_IndiceProjet=" & IdIndiceProjet & "  "
Sql = Sql & "AND Composants_Ecart.NUMCOMP Is Null;"

Set Rs = Con.OpenRecordSet(Sql)
Set MySheet = IsertSheet(MyWorkbook, "Composants_Ecart")
Set MyRange = MySheet.Cells(1, 1).CurrentRegion
L = MyRange.Rows.Count
If L > 1 Then L = L + 1
If Rs.EOF = False Then
MajEcartComposants = True
MyRange(L, 1) = "Enregistrement Ajouter"
FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L, Rs.Fields.Count).Address), 40, True, True, xlCenter, xlCenter


Set MyRange = MySheet.Cells(1, 1).CurrentRegion
L = MyRange.Rows.Count + 1

For I = 0 To Rs.Fields.Count - 1
    MyRange(L, I + 1) = Rs(I).Name
Next
FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L, Rs.Fields.Count).Address), 15, False, True, xlCenter, xlCenter

While Rs.EOF = False
L = L + 1
   For I = 0 To Rs.Fields.Count - 1
    MyRange(L, I + 1) = "" & Rs(I).Value
    Next
    FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L, Rs.Fields.Count).Address), 2, False, True, xlCenter, xlCenter

    Rs.MoveNext
Wend



End If
Set Rs = Con.CloseRecordSet(Rs)
Sql = "SELECT Composants.*,Composants_Ecart.* "
Sql = Sql & "FROM Composants INNER JOIN Composants_Ecart  "
Sql = Sql & "ON (Composants.NUMCOMP = Composants_Ecart.NUMCOMP)  "
Sql = Sql & "AND (Composants.Id_IndiceProjet = Composants_Ecart.Id_IndiceProjet) "
Sql = Sql & "WHERE Composants.Id_IndiceProjet=" & IdIndiceProjet & "  "


Set Rs = Con.OpenRecordSet(Sql)
Set MySheet = IsertSheet(MyWorkbook, "Composants_Ecart")
Set MyRange = MySheet.Cells(1, 1).CurrentRegion
L = MyRange.Rows.Count
If L > 1 Then L = L + 1
If Rs.EOF = False Then
MyRange(L, 1) = "Enregistrement Modifier"
FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L, (Rs.Fields.Count / 2)).Address), 40, True, True, xlCenter, xlCenter


Set MyRange = MySheet.Cells(1, 1).CurrentRegion
L = MyRange.Rows.Count + 1
For I = 0 To (Rs.Fields.Count / 2) - 1
    MyRange(L, I + 1) = Rs(I).Name
Next
FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L, Rs.Fields.Count / 2).Address), 15, False, True, xlCenter, xlCenter

While Rs.EOF = False
L = L + 1
modifire = False
   For I = 0 To (Rs.Fields.Count - 2) / 2
    If "" & Rs(I).Value <> "" & Rs(6 + I).Value Then
        If I > 0 Then
        MajEcartComposants = True
        modifire = True
        Exit For
        End If
    End If
   
    Next
    If modifire = True Then
   For I = 0 To (Rs.Fields.Count - 2) / 2
    MyRange(L, I + 1) = "" & Rs(I).Value & Chr(10) & "" & Rs(6 + I).Value
    If "" & Rs(I).Value <> "" & Rs(6 + I).Value And I > 0 Then
        MyRange(L, I + 1).Font.ColorIndex = 3
       
    End If
    
    Next
          FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L, Rs.Fields.Count / 2).Address), 2, False, True, xlCenter, xlCenter
    Else
        L = L - 1
    End If
  

    Rs.MoveNext
Wend
End If
Set Rs = Con.CloseRecordSet(Rs)
Set MyRange = MyRange(1, 1).CurrentRegion
If MyRange.Rows.Count = 1 Then Exit Function

Sql = "SELECT  T_indiceProjet.Ensemble, T_indiceProjet.Equipement,[Li] & '_' & [LI_Indice] AS Liste, "
Sql = Sql & "[PI] & '_' & [PI_Indice] AS Piece, T_indiceProjet.Client, T_indiceProjet.CleAc "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"
Dim RsEntetePage As Recordset
Set RsEntetePage = Con.OpenRecordSet(Sql)
Set MyRange = MyRange(1, 1).CurrentRegion
If PortraitPaysage = 0 Then PortraitPaysage = 2
MiseEnPage MySheet, MyRange, "Affaire : " & RsEntetePage!CleAc & Chr(10) & _
    "Câblage : " & Replace("" & RsEntetePage!Ensemble, vbCrLf, " ") & Chr(10) & _
     "Pièce : " & RsEntetePage!Piece, vbCrLf & "Equipement : " & Replace("" & RsEntetePage!Equipement, vbCrLf, " ") & Chr(10) & _
     "Liste : " & RsEntetePage!Liste _
    , "ENCELADE" & Chr(10) & "Client : " & RsEntetePage!Client & Chr(10) & Format(Date, "dd-mmm-yyyy"), "", "&P/&N", "", 80, "", False, PortraitPaysage, False, True, True

    MaJEncadreXls MyRange, xlThin, xlThin, xlHairline, xlHairline

Set RsEntetePage = Con.CloseRecordSet(RsEntetePage)
End Function


Function MajEcartExcel(IdIndiceProjet As Long, MyWorkbook As EXCEL.Workbook, L As Long, C As Long, RsSuprimer As Recordset, RsAjouter As Recordset, RsModifier As Recordset, SheetName As String, Optional Txt = "") As Boolean
Dim Sql As String
Dim TruveSheet As Boolean
Dim boolTxt As Boolean
Dim MySheet As EXCEL.Worksheet
Dim MyRange As Range
'MyWorkbook.Application.Visible = True

RecherModifier RsModifier
If RsModifier.EOF = False Then
   
   
    Set MySheet = IsertSheet(MyWorkbook, SheetName)
      MajEcartExcel = True
     Set MyRange = MySheet.Cells(1, 1).CurrentRegion
    L = MyRange.Rows.Count
    If L > 1 Then L = L + 1
    If Trim("" & Txt) <> "" And boolTxt = False Then
        boolTxt = True
        T_Txt = Split(Txt, Chr(10))
        I2 = 0
        For I = LBound(T_Txt) To UBound(T_Txt)
            If Trim("" & T_Txt(I)) <> "" Then
                MyRange(L + I - I2, 1) = T_Txt(I)
                FormatExcelPlage MySheet.Range(MyRange(L + I - I2, 1).Address & ":" & MyRange(L + I - I2, RsModifier.Fields.Count / 2).Address), 2, True, False, xlCenter, xlCenter
    
            Else
                I2 = I2 + 1
            End If
        Next
    
           Set MyRange = MySheet.Cells(1, 1).CurrentRegion
            L = MyRange.Rows.Count
            If L > 1 Then L = L + 1
    End If
End If
If RsSuprimer.EOF = False Then
    MajEcartExcel = True
     
    Set MySheet = IsertSheet(MyWorkbook, SheetName)
    

    Set MyRange = MySheet.Cells(1, 1).CurrentRegion
    L = MyRange.Rows.Count
    If L > 1 Then L = L + 1
   
    If Trim("" & Txt) <> "" And boolTxt = False Then
        boolTxt = True
        T_Txt = Split(Txt, Chr(10))
        I2 = 0
        For I = LBound(T_Txt) To UBound(T_Txt)
            If Trim("" & T_Txt(I)) <> "" Then
                MyRange(L + I - I2, 1) = T_Txt(I)
                FormatExcelPlage MySheet.Range(MyRange(L + I - I2, 1).Address & ":" & MyRange(L + I - I2, RsModifier.Fields.Count / 2).Address), 2, True, False, xlCenter, xlCenter
    
            Else
                I2 = I2 + 1
            End If
        Next
    
           Set MyRange = MySheet.Cells(1, 1).CurrentRegion
            L = MyRange.Rows.Count
            If L > 1 Then L = L + 1
    End If
     MyRange(MyRange.Rows.Count, 1) = "Enregistrement Suprimer"
      FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L, RsAjouter.Fields.Count).Address), 40, True, True, xlCenter, xlCenter
    L = L + 1
    For I = 0 To RsSuprimer.Fields.Count - 1
        MyRange(L, I + 1) = RsSuprimer(I).Name
    Next
    FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L, RsSuprimer.Fields.Count).Address), 15, False, True, xlCenter, xlCenter

    While RsSuprimer.EOF = False
        DoEvents
        L = L + 1
        For I = 0 To RsSuprimer.Fields.Count - 1
            MyRange(L, I + 1) = "" & RsSuprimer(I).Value
        Next
            FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L, RsSuprimer.Fields.Count).Address), 2, False, True, xlCenter, xlCenter

            RsSuprimer.MoveNext
    Wend

End If
If RsAjouter.EOF = False Then
MajEcartExcel = True


    Set MySheet = IsertSheet(MyWorkbook, SheetName)
    Set MyRange = MySheet.Cells(1, 1).CurrentRegion
    L = MyRange.Rows.Count
    If L > 1 Then L = L + 1
     If Trim("" & Txt) <> "" And boolTxt = False Then
        boolTxt = True
        T_Txt = Split(Txt, Chr(10))
        I2 = 0
        For I = LBound(T_Txt) To UBound(T_Txt)
            If Trim("" & T_Txt(I)) <> "" Then
                MyRange(L + I - I2, 1) = T_Txt(I)
                FormatExcelPlage MySheet.Range(MyRange(L + I - I2, 1).Address & ":" & MyRange(L + I - I2, RsModifier.Fields.Count / 2).Address), 2, True, False, xlCenter, xlCenter
    
            Else
                I2 = I2 + 1
            End If
        Next
    
           Set MyRange = MySheet.Cells(1, 1).CurrentRegion
            L = MyRange.Rows.Count
            If L > 1 Then L = L + 1
    End If
    If RsAjouter.EOF = False Then
    DoEvents
    MajEcartExcel = True
    
    
    
    If Trim("" & Txt) <> "" And boolTxt = False Then
        boolTxt = True
        T_Txt = Split(Txt, Chr(10))
        I2 = 0
        For I = LBound(T_Txt) To UBound(T_Txt)
            If Trim("" & T_Txt(I)) <> "" Then
                MyRange(L + I - I2, 1) = T_Txt(I)
                FormatExcelPlage MySheet.Range(MyRange(L + I - I2, 1).Address & ":" & MyRange(L + I - I2, RsModifier.Fields.Count / 2).Address), 2, True, False, xlCenter, xlCenter
    
            Else
                I2 = I2 + 1
            End If
        Next
    
        L = L + 1
    End If
    MyRange(L, 1) = "Enregistrement Ajouter"
    FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L, RsAjouter.Fields.Count).Address), 40, True, True, xlCenter, xlCenter

L = L + 1
    For I = 0 To RsAjouter.Fields.Count - 1
        MyRange(L, I + 1) = RsAjouter(I).Name
    Next
    FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L, RsAjouter.Fields.Count).Address), 15, False, True, xlCenter, xlCenter

    While RsAjouter.EOF = False
        DoEvents
        L = L + 1
        For I = 0 To RsAjouter.Fields.Count - 1
        MyRange(L, I + 1) = "" & RsAjouter(I).Value
        Next
        FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L, RsAjouter.Fields.Count).Address), 2, False, True, xlCenter, xlCenter

        RsAjouter.MoveNext
    Wend
End If


End If


If RsModifier.EOF = False Then
MajEcartExcel = True
RecherModifier RsModifier
Set MySheet = IsertSheet(MyWorkbook, SheetName)
Set MyRange = MySheet.Cells(1, 1).CurrentRegion
L = MyRange.Rows.Count
If L > 1 Then L = L + 1


Set MyRange = MySheet.Cells(1, 1).CurrentRegion
L = MyRange.Rows.Count
If L > 1 Then L = L + 1
If Trim("" & Txt) <> "" And boolTxt = False Then
        boolTxt = True
        T_Txt = Split(Txt, Chr(10))
        I2 = 0
        For I = LBound(T_Txt) To UBound(T_Txt)
            If Trim("" & T_Txt(I)) <> "" Then
                MyRange(L + I - I2, 1) = T_Txt(I)
                FormatExcelPlage MySheet.Range(MyRange(L + I - I2, 1).Address & ":" & MyRange(L + I - I2, RsModifier.Fields.Count / 2).Address), 2, True, False, xlCenter, xlCenter
    
            Else
                I2 = I2 + 1
            End If
        Next
    
        L = L + 1
    End If
    MyRange(L, 1) = "Enregistrement Modifier"
FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L, (RsModifier.Fields.Count / 2)).Address), 40, True, True, xlCenter, xlCenter
L = L + 1
For I = 0 To (RsModifier.Fields.Count / 2) - 1
    MyRange(L, I + 1) = RsModifier(I).Name
Next
FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L, RsModifier.Fields.Count / 2).Address), 15, False, True, xlCenter, xlCenter

While RsModifier.EOF = False
L = L + 1
modifire = False
   For I = 0 To (RsModifier.Fields.Count - 2) / 2
    
    If UCase(Trim("" & RsModifier(I).Value)) <> UCase(Trim("" & RsModifier(((RsModifier.Fields.Count) / 2) + I).Value)) Then
       
        modifire = True
        MajEcartExcel = True
        Exit For
        
    End If
   
    Next
    If modifire = True Then
   For I = 0 To (RsModifier.Fields.Count - 2) / 2
   a = RsModifier(I).Name
   aa = RsModifier(((RsModifier.Fields.Count)) / 2 + I).Name
    If UCase(RsModifier.Fields(I).Name) = UCase("Avant/Après") Then
        MyRange(L, I + 1) = "Avant"
        MyRange(L + 1, I + 1) = "Après"
    End If
   If RsModifier(I).Type = adBoolean Then
        If RsModifier(I).Value = True Then
             MyRange(L, I + 1) = "1"  '& RsModifier(((RsModifier.Fields.Count)) / 2 + i).Value
        Else
            MyRange(L, I + 1) = "0"   '& RsModifier(((RsModifier.Fields.Count)) / 2 + i).Value
        End If
        If RsModifier(((RsModifier.Fields.Count)) / 2 + I) = True Then
             MyRange(L + 1, I + 1) = "1" '& RsModifier(((RsModifier.Fields.Count)) / 2 + i).Value
        Else
            MyRange(L + 1, I + 1) = "0" '& RsModifier(((RsModifier.Fields.Count)) / 2 + i).Value
        End If
   Else
   a = RsModifier(I).Name
   aa = RsModifier(((RsModifier.Fields.Count)) / 2 + I).Name
    If UCase(RsModifier.Fields(I).Name) <> UCase("Avant/Après") Then
        MyRange(L, I + 1) = "" & RsModifier(I).Value
        MyRange(L + 1, I + 1) = "" & RsModifier(((RsModifier.Fields.Count)) / 2 + I).Value
    End If
    End If
    If "" & RsModifier(I).Value <> "" & RsModifier(((RsModifier.Fields.Count)) / 2 + I).Value Then
        MyRange(L + 1, I + 1).Font.ColorIndex = 3
       
    End If
    
    Next
          FormatExcelPlage MySheet.Range(MyRange(L, 1).Address & ":" & MyRange(L + 1, RsModifier.Fields.Count / 2).Address), 2, False, True, xlCenter, xlCenter
    Else
        
    End If
  Set MyRange = MySheet.Cells(1, 1).CurrentRegion
    L = MyRange.Rows.Count

    RsModifier.MoveNext
Wend
End If
If MajEcartExcel = True Then
'Set Rs = Con.CloseRecordSet(Rs)
Set MyRange = MyRange(1, 1).CurrentRegion
If MyRange.Rows.Count = 1 Then Exit Function

Sql = "SELECT  T_indiceProjet.Ensemble, T_indiceProjet.Equipement,[Li] & '_' & [LI_Indice] AS Liste, "
Sql = Sql & "[PI] & '_' & [PI_Indice] AS Piece, T_indiceProjet.Client, T_indiceProjet.CleAc "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"
Dim RsEntetePage As Recordset
Set RsEntetePage = Con.OpenRecordSet(Sql)
Set MyRange = MyRange(1, 1).CurrentRegion
If PortraitPaysage = 0 Then PortraitPaysage = 2
MiseEnPage MySheet, MyRange, "Affaire : " & RsEntetePage!CleAc & Chr(10) & _
    "Câblage : " & Replace("" & RsEntetePage!Ensemble, vbCrLf, " ") & Chr(10) & _
     "Pièce : " & RsEntetePage!Piece, vbCrLf & "Equipement : " & Replace("" & RsEntetePage!Equipement, vbCrLf, " ") & Chr(10) & _
     "Liste : " & RsEntetePage!Liste _
    , "ENCELADE" & Chr(10) & "Client : " & RsEntetePage!Client & Chr(10) & Format(Date, "dd-mmm-yyyy"), "", "&P/&N", "", 80, "", False, PortraitPaysage, False, True, True

    MaJEncadreXls MyRange, xlThin, xlThin, xlHairline, xlHairline

Set RsEntetePage = Con.CloseRecordSet(RsEntetePage)
End If
End Function

Sub RecherModifier(RsModifier As Recordset)

While RsModifier.EOF = False
DoEvents
   For I = 0 To (RsModifier.Fields.Count - 2) / 2
    If UCase(Trim("" & RsModifier(I).Value)) <> UCase(Trim("" & RsModifier(((RsModifier.Fields.Count) / 2) + I).Value)) Then
       
        Exit Sub
        
    End If
   
    Next
    RsModifier.MoveNext
Wend

End Sub


Public Function RetourneNbRows(MyRange As Object) As Long
RetourneNbRows = MyRange.Rows.Count - 1
If RetourneNbRows = 0 Then RetourneNbRows = 1
End Function
