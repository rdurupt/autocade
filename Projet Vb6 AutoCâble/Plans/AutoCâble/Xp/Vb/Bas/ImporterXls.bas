Attribute VB_Name = "ImporterXls"
Public Sub ImporteXls(Xls As String, IdIndiceProjet As Long)
Dim Sql As String
Dim Rs As Recordset
Dim IdProjet As Long
Dim IdIndice As Long

DoEvents

Set TableauPath = funPath
IdIndice = IdIndiceProjet

Dim MyEcel As New EXCEL.Application
Dim MyClasseur As EXCEL.Workbook
Dim MySheet As EXCEL.Worksheet
Dim Myrange As EXCEL.Range
Set MyClasseur = MyEcel.Workbooks.Open(Xls)
'***********************************************************************************************************************
'*                                          Supprime le contenu des tables temporaire.                                 *
Sql = "DELETE Xls_Critères.*  FROM Xls_Critères "
Sql = Sql & "where Xls_Critères.Job=" & NmJob & ";"
Con.Exequte Sql


Sql = "DELETE xls_Ligne_Tableau_fils.*  FROM xls_Ligne_Tableau_fils "
Sql = Sql & "where xls_Ligne_Tableau_fils.Job=" & NmJob & ";"
Con.Exequte Sql

Sql = "DELETE Xls_Connecteurs.* FROM Xls_Connecteurs "
Sql = Sql & "WHERE Xls_Connecteurs.Job=" & NmJob & ";"
Con.Exequte Sql

Sql = "DELETE Xls_Nota.* FROM Xls_Nota "
Sql = Sql & "WHERE Xls_Nota.Job=" & NmJob & ";"
Con.Exequte Sql

Sql = "DELETE Xls_Composants.* FROM Xls_Composants "
Sql = Sql & "WHERE Xls_Composants.Job=" & NmJob & ";"
Con.Exequte Sql

Sql = "DELETE Xls_Noeuds.* FROM Xls_Noeuds "
Sql = Sql & "WHERE Xls_Noeuds.Job=" & NmJob & ";"
Con.Exequte Sql
'MyEcel.Visible = True

'***********************************************************************************************************************
'*                                          Supprime le contenu des tables Ecart.                                     *
Sql = "DELETE T_Critères_Ecart.*  FROM T_Critères_Ecart "
Sql = Sql & "where T_Critères_Ecart.Id_IndiceProjet=" & IdIndiceProjet & ";"
Con.Exequte Sql


Sql = "DELETE Ligne_Tableau_fils_Ecart.*  FROM Ligne_Tableau_fils_Ecart "
Sql = Sql & "where Ligne_Tableau_fils_Ecart.Id_IndiceProjet=" & IdIndiceProjet & ";"
Con.Exequte Sql

Sql = "DELETE Connecteurs_Ecart.* FROM Connecteurs_Ecart "
Sql = Sql & "WHERE Connecteurs_Ecart.Id_IndiceProjet=" & IdIndiceProjet & ";"
Con.Exequte Sql

Sql = "DELETE Nota_Ecart.* FROM Nota_Ecart "
Sql = Sql & "WHERE Nota_Ecart.Id_IndiceProjet=" & IdIndiceProjet & ";"
Con.Exequte Sql

Sql = "DELETE Composants_Ecart.* FROM Composants_Ecart "
Sql = Sql & "WHERE Composants_Ecart.Id_IndiceProjet=" & IdIndiceProjet & ";"
Con.Exequte Sql

Sql = "DELETE T_Noeuds_Ecart.* FROM T_Noeuds_Ecart "
Sql = Sql & "WHERE T_Noeuds_Ecart.Id_IndiceProjet=" & IdIndiceProjet & ";"
Con.Exequte Sql

'***********************************************************************************************************************
'*                                        Sauvegarde les anciennes valeurs                                            *

Sql = "INSERT INTO T_Critères_Ecart SELECT T_Critères.* FROM T_Critères "
Sql = Sql & "WHERE T_Critères.Id_IndiceProjet=" & IdIndiceProjet & ";"
Con.Exequte Sql

Sql = "INSERT INTO Connecteurs_Ecart SELECT Connecteurs.* FROM Connecteurs "
Sql = Sql & "WHERE Connecteurs.Id_IndiceProjet=" & IdIndiceProjet & ";"
Con.Exequte Sql

Sql = "INSERT INTO Nota_Ecart SELECT Nota.* FROM Nota "
Sql = Sql & "WHERE Nota.Id_IndiceProjet=" & IdIndiceProjet & ";"
Con.Exequte Sql

Sql = "INSERT INTO Composants_Ecart SELECT Composants.* FROM Composants "
Sql = Sql & "WHERE Composants.Id_IndiceProjet=" & IdIndiceProjet & ";"
Con.Exequte Sql

Sql = "INSERT INTO Ligne_Tableau_fils_Ecart SELECT Ligne_Tableau_fils.* FROM Ligne_Tableau_fils "
Sql = Sql & "WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & IdIndiceProjet & ";"
Con.Exequte Sql

Sql = "INSERT INTO T_Noeuds_Ecart SELECT T_Noeuds.* FROM T_Noeuds "
Sql = Sql & "WHERE T_Noeuds.Id_IndiceProjet=" & IdIndiceProjet & ";"
Con.Exequte Sql

'***********************************************************************************************************************
'*                              Importe la liste des Noeuds dans la table temporaire:                                    *

Set MySheet = MyClasseur.Worksheets("Noeuds")
'MySheet.Application.Visible = True
Set Myrange = MySheet.Range("A1").CurrentRegion
 FormBarGrah.ProgressBar1.Value = 0
 FormBarGrah.ProgressBar1.Max = Myrange.Rows.Count
 FormBarGrah.ProgressBar1Caption.Caption = " Importe la liste des Noeuds "

For Row = 2 To Myrange.Rows.Count
 IncremanteBarGrah FormBarGrah
DoEvents

   Sql = sqlRange(Myrange, Row, "Xls_Noeuds")
   Con.Exequte Sql
Next Row

'***********************************************************************************************************************
'*                              Importe la liste des Critères dans la table temporaire:                                    *

Set MySheet = MyClasseur.Worksheets("Critères")

Set Myrange = MySheet.Range("A1").CurrentRegion
 FormBarGrah.ProgressBar1.Value = 0
 FormBarGrah.ProgressBar1.Max = Myrange.Rows.Count
 FormBarGrah.ProgressBar1Caption.Caption = " Importe la liste de Critères "

For Row = 2 To Myrange.Rows.Count
 IncremanteBarGrah FormBarGrah
DoEvents

   Sql = sqlRange(Myrange, Row, "Xls_Critères")
   Con.Exequte Sql
Next Row

'***********************************************************************************************************************
'*                              Importe la liste des fils dans la table temporaire:                                    *

Set MySheet = MyClasseur.Worksheets("Ligne_Tableau_fils")

Set Myrange = MySheet.Range("A1").CurrentRegion
 FormBarGrah.ProgressBar1.Value = 0
 FormBarGrah.ProgressBar1.Max = Myrange.Rows.Count
 FormBarGrah.ProgressBar1Caption.Caption = " Importe la liste de fils"

For Row = 2 To Myrange.Rows.Count
 IncremanteBarGrah FormBarGrah
DoEvents

   Sql = sqlRange(Myrange, Row, "Xls_Ligne_Tableau_fils")
   Con.Exequte Sql
Next Row

'***********************************************************************************************************************
'*                              Importe la liste des Connecteurs dans la table temporaire:                             *

Set MySheet = MyClasseur.Worksheets("Connecteurs")
Set Myrange = MySheet.Range("A1").CurrentRegion
 
 FormBarGrah.ProgressBar1.Value = 0
 FormBarGrah.ProgressBar1.Max = Myrange.Rows.Count
 FormBarGrah.ProgressBar1Caption.Caption = " Importe la liste des Connecteurs"

For Row = 2 To Myrange.Rows.Count
 IncremanteBarGrah FormBarGrah
   Sql = sqlRange(Myrange, Row, "Xls_Connecteurs")
   Con.Exequte Sql
Next Row

'***********************************************************************************************************************
'*                              Importe la liste des Composants dans la table temporaire:                              *

Set MySheet = MyClasseur.Worksheets("Composants")
Set Myrange = MySheet.Range("A1").CurrentRegion

 FormBarGrah.ProgressBar1.Value = 0
 FormBarGrah.ProgressBar1.Max = Myrange.Rows.Count
 FormBarGrah.ProgressBar1Caption.Caption = " Importe la liste des Composants"

For Row = 2 To Myrange.Rows.Count
 IncremanteBarGrah FormBarGrah
   Sql = sqlRange(Myrange, Row, "Xls_Composants")
   Con.Exequte Sql
Next Row

'***********************************************************************************************************************
'*                              Importe la liste des Notas dans la table temporaire:                                   *

Set MySheet = MyClasseur.Worksheets("Notas")
Set Myrange = MySheet.Range("A1").CurrentRegion
 
 FormBarGrah.ProgressBar1.Value = 0
 FormBarGrah.ProgressBar1.Max = Myrange.Rows.Count
 FormBarGrah.ProgressBar1Caption.Caption = " Importe la liste des Notas"

For Row = 2 To Myrange.Rows.Count
 IncremanteBarGrah FormBarGrah
   Sql = sqlRange(Myrange, Row, "Xls_Nota")
   Con.Exequte Sql
Next Row

Set Myrange = Nothing
Set MySheet = Nothing
MyClasseur.Close False
Set MyClasseur = Nothing
MyEcel.Quit
Set MyEcel = Nothing


MajBase IdIndice
MajEcart IdIndiceProjet, MyEcel
'***********************************************************************************************************************
'*                                          Supprime le contenu des tables Ecart.                                     *
Sql = "DELETE T_Critères_Ecart.*  FROM T_Critères_Ecart "
Sql = Sql & "where T_Critères_Ecart.Id_IndiceProjet=" & IdIndiceProjet & ";"
Con.Exequte Sql


Sql = "DELETE Ligne_Tableau_fils_Ecart.*  FROM Ligne_Tableau_fils_Ecart "
Sql = Sql & "where Ligne_Tableau_fils_Ecart.Id_IndiceProjet=" & IdIndiceProjet & ";"
Con.Exequte Sql

Sql = "DELETE Connecteurs_Ecart.* FROM Connecteurs_Ecart "
Sql = Sql & "WHERE Connecteurs_Ecart.Id_IndiceProjet=" & IdIndiceProjet & ";"
Con.Exequte Sql

Sql = "DELETE Nota_Ecart.* FROM Nota_Ecart "
Sql = Sql & "WHERE Nota_Ecart.Id_IndiceProjet=" & IdIndiceProjet & ";"
Con.Exequte Sql

Sql = "DELETE Composants_Ecart.* FROM Composants_Ecart "
Sql = Sql & "WHERE Composants_Ecart.Id_IndiceProjet=" & IdIndiceProjet & ";"
Con.Exequte Sql

Sql = "DELETE T_Noeuds_Ecart.* FROM T_Noeuds_Ecart "
Sql = Sql & "WHERE T_Noeuds_Ecart.Id_IndiceProjet=" & IdIndiceProjet & ";"
Con.Exequte Sql
'***********************************************************************************************************************
'*                                          Supprime le contenu des tables temporaire.                                 *

Sql = "DELETE xls_Ligne_Tableau_fils.*  FROM xls_Ligne_Tableau_fils "
Sql = Sql & "where xls_Ligne_Tableau_fils.Job=" & NmJob & ";"
Con.Exequte Sql

Sql = "DELETE Xls_Connecteurs.* FROM Xls_Connecteurs "
Sql = Sql & "WHERE Xls_Connecteurs.Job=" & NmJob & ";"
Con.Exequte Sql

Sql = "DELETE Xls_Nota.* FROM Xls_Nota "
Sql = Sql & "WHERE Xls_Nota.Job=" & NmJob & ";"
Con.Exequte Sql

Sql = "DELETE Xls_Composants.* FROM Xls_Composants "
Sql = Sql & "WHERE Xls_Composants.Job=" & NmJob & ";"
Con.Exequte Sql

 FormBarGrah.ProgressBar1.Value = 0
 '***********************************************************************************************************************

 FormBarGrah.ProgressBar1Caption.Caption = " Fin du traitement"
End Sub

Sub MajEcart(IdIndiceProjet As Long, MyEcel As EXCEL.Application)
 Set TableauPath = funPath
Dim L As Long
Dim C As Long
Dim boolSave As Boolean
Dim Sql As String
Dim RsSuprimer As Recordset
Dim RsAjouter As Recordset
Dim RsModifier As Recordset
boolSave = False
   PathArchiveAutocad = TableauPath.Item("PathArchiveAutocad")
     If Left(PathArchiveAutocad, 2) <> "\\" And Left(PathArchiveAutocad, 1) = "\" Then PathArchiveAutocad = TableauPath.Item("PathServer") & PathArchiveAutocad
  If Right(PathArchiveAutocad, 2) = "\\" Then PathArchiveAutocad = Mid(PathArchiveAutocad, 1, Len(MyPath) - 1)
'MyEcel.Visible = True
Set MyWorkbook = MyEcel.Workbooks.Add


Sql = "SELECT Nota_Ecart.ACTIVER,Nota_Ecart.NOTA, Nota_Ecart.NUMNOTA "
Sql = Sql & "FROM Nota_Ecart LEFT JOIN Nota ON (Nota_Ecart.NUMNOTA = Nota.NUMNOTA) AND (Nota_Ecart.Id_IndiceProjet = Nota.Id_IndiceProjet) "
Sql = Sql & "WHERE Nota_Ecart.Id_IndiceProjet=" & IdIndiceProjet & " "
Sql = Sql & "AND Nota.NUMNOTA Is Null;"
Debug.Print Sql


Set RsSuprimer = Con.OpenRecordSet(Sql)

Sql = "SELECT Nota.ACTIVER,Nota.NOTA, Nota.NUMNOTA "
Sql = Sql & "FROM Nota LEFT JOIN Nota_Ecart  "
Sql = Sql & "ON (Nota.NUMNOTA = Nota_Ecart.NUMNOTA)  "
Sql = Sql & "AND (Nota.Id_IndiceProjet = Nota_Ecart.Id_IndiceProjet) "
Sql = Sql & "WHERE Nota.Id_IndiceProjet=" & IdIndiceProjet & "  "
Sql = Sql & "AND Nota_Ecart.NUMNOTA Is Null;"
Debug.Print Sql

Set RsAjouter = Con.OpenRecordSet(Sql)


Sql = "SELECT '' AS [Avant/Après],Nota_Ecart.ACTIVER, Nota_Ecart.NOTA, Nota_Ecart.NUMNOTA, '' AS Expr1,Nota.ACTIVER, Nota.NOTA, Nota.NUMNOTA "
Sql = Sql & "FROM Nota INNER JOIN Nota_Ecart  "
Sql = Sql & "ON (Nota.NUMNOTA = Nota_Ecart.NUMNOTA)  "
Sql = Sql & "AND (Nota.Id_IndiceProjet = Nota_Ecart.Id_IndiceProjet) "
Sql = Sql & "WHERE Nota_Ecart.Id_IndiceProjet=" & IdIndiceProjet & ";  "
Debug.Print Sql


Set RsModifier = Con.OpenRecordSet(Sql)



L = 0: C = 0
If MajEcartExcel(IdIndiceProjet, MyWorkbook, L, C, RsSuprimer, RsAjouter, RsModifier, "Notas_Ecart") = True Then boolSave = True

Sql = "SELECT Composants_Ecart.ACTIVER, Composants_Ecart.DESIGNCOMP, Composants_Ecart.NUMCOMP, Composants_Ecart.REFCOMP, Composants_Ecart.Path "
Sql = Sql & "FROM Composants_Ecart LEFT JOIN Composants  "
Sql = Sql & "ON (Composants_Ecart.NUMCOMP = Composants.NUMCOMP)  "
Sql = Sql & "AND (Composants_Ecart.Id_IndiceProjet = Composants.Id_IndiceProjet) "
Sql = Sql & "WHERE Composants.NUMCOMP Is Null "
Sql = Sql & "AND Composants_Ecart.Id_IndiceProjet=" & IdIndiceProjet & " ;"
Debug.Print Sql

Set RsSuprimer = Con.OpenRecordSet(Sql)

Sql = "SELECT  Composants.ACTIVER,Composants.DESIGNCOMP, Composants.NUMCOMP, Composants.REFCOMP, Composants.Path "
Sql = Sql & "FROM Composants LEFT JOIN Composants_Ecart  "
Sql = Sql & "ON (Composants.NUMCOMP = Composants_Ecart.NUMCOMP)  "
Sql = Sql & "AND (Composants.Id_IndiceProjet = Composants_Ecart.Id_IndiceProjet) "
Sql = Sql & "WHERE Composants.Id_IndiceProjet=" & IdIndiceProjet & "  "
Sql = Sql & "AND Composants_Ecart.NUMCOMP Is Null;"
Debug.Print Sql
Set RsAjouter = Con.OpenRecordSet(Sql)

Sql = "SELECT '' AS [Avant/Après],Composants_Ecart.ACTIVER, Composants_Ecart.DESIGNCOMP, Composants_Ecart.NUMCOMP,   "
Sql = Sql & "Composants_Ecart.REFCOMP, Composants_Ecart.Path, '' AS Expr1,Composants.ACTIVER, Composants.DESIGNCOMP,   "
Sql = Sql & "Composants.NUMCOMP, Composants.REFCOMP, Composants.Path  "
Sql = Sql & "FROM Composants INNER JOIN Composants_Ecart  "
Sql = Sql & "ON (Composants.NUMCOMP = Composants_Ecart.NUMCOMP)  "
Sql = Sql & "AND (Composants.Id_IndiceProjet = Composants_Ecart.Id_IndiceProjet) "
Sql = Sql & "WHERE Composants.Id_IndiceProjet=" & IdIndiceProjet & "  "
Debug.Print Sql

Set RsModifier = Con.OpenRecordSet(Sql)

L = 0: C = 0
If MajEcartExcel(IdIndiceProjet, MyWorkbook, L, C, RsSuprimer, RsAjouter, RsModifier, "Composants_Ecart") = True Then boolSave = True




Sql = "SELECT  T_Noeuds_Ecart.ACTIVER,T_Noeuds_Ecart.Fleche_Droite,T_Noeuds_Ecart.TORON_PRINCIPAL,  "
Sql = Sql & "T_Noeuds_Ecart.NŒUDS, T_Noeuds_Ecart.LONGUEUR, T_Noeuds_Ecart.LONGUEUR_CUMULEE, "
Sql = Sql & "T_Noeuds_Ecart.DESIGN_HAB, T_Noeuds_Ecart.CODE_RSA, T_Noeuds_Ecart.CODE_PSA, "
Sql = Sql & "T_Noeuds_Ecart.CODE_ENC, T_Noeuds_Ecart.DIAMETRE, T_Noeuds_Ecart.CLASSE_T "
Sql = Sql & "FROM T_Noeuds_Ecart LEFT JOIN T_Noeuds "
Sql = Sql & "ON (T_Noeuds_Ecart.NŒUDS = T_Noeuds.NŒUDS) "
Sql = Sql & "AND (T_Noeuds_Ecart.Id_IndiceProjet = T_Noeuds.Id_IndiceProjet)"
Sql = Sql & "WHERE T_Noeuds.NŒUDS Is Null "
Sql = Sql & "AND T_Noeuds_Ecart.Id_IndiceProjet=" & IdIndiceProjet & " ;"
Debug.Print Sql

Set RsSuprimer = Con.OpenRecordSet(Sql)

Sql = "SELECT T_Noeuds.ACTIVER, T_Noeuds.Fleche_Droite, T_Noeuds.TORON_PRINCIPAL,  "
Sql = Sql & "T_Noeuds.NŒUDS, T_Noeuds.LONGUEUR, T_Noeuds.LONGUEUR_CUMULEE, T_Noeuds.DESIGN_HAB,  "
Sql = Sql & "T_Noeuds.CODE_RSA, T_Noeuds.CODE_PSA, T_Noeuds.CODE_ENC, T_Noeuds.DIAMETRE, T_Noeuds.CLASSE_T "
Sql = Sql & "FROM T_Noeuds LEFT JOIN T_Noeuds_Ecart  "
Sql = Sql & "ON (T_Noeuds.Id_IndiceProjet = T_Noeuds_Ecart.Id_IndiceProjet)  "
Sql = Sql & "AND (T_Noeuds.NŒUDS = T_Noeuds_Ecart.NŒUDS) "
Sql = Sql & "WHERE T_Noeuds.Id_IndiceProjet=" & IdIndiceProjet & "  "
Sql = Sql & "AND T_Noeuds_Ecart.NŒUDS Is Null;"
Debug.Print Sql
Set RsAjouter = Con.OpenRecordSet(Sql)

Sql = "SELECT '' AS [Avant/Après],T_Noeuds_Ecart.ACTIVER, T_Noeuds_Ecart.Fleche_Droite, T_Noeuds_Ecart.TORON_PRINCIPAL,  "
Sql = Sql & " T_Noeuds_Ecart.NŒUDS, T_Noeuds_Ecart.LONGUEUR,  "
Sql = Sql & "T_Noeuds_Ecart.LONGUEUR_CUMULEE, T_Noeuds_Ecart.DESIGN_HAB, T_Noeuds_Ecart.CODE_RSA,  "
Sql = Sql & "T_Noeuds_Ecart.CODE_PSA, T_Noeuds_Ecart.CODE_ENC, T_Noeuds_Ecart.DIAMETRE,  "
Sql = Sql & "T_Noeuds_Ecart.CLASSE_T, '' AS Expr1,T_Noeuds.ACTIVER, T_Noeuds.Fleche_Droite, T_Noeuds.TORON_PRINCIPAL,  "
Sql = Sql & " T_Noeuds.NŒUDS, T_Noeuds.LONGUEUR, T_Noeuds.LONGUEUR_CUMULEE,  "
Sql = Sql & "T_Noeuds.DESIGN_HAB, T_Noeuds.CODE_RSA, T_Noeuds.CODE_PSA, T_Noeuds.CODE_ENC,  "
Sql = Sql & "T_Noeuds.DIAMETRE, T_Noeuds.CLASSE_T "
Sql = Sql & "FROM T_Noeuds INNER JOIN T_Noeuds_Ecart  "
Sql = Sql & "ON (T_Noeuds.NŒUDS = T_Noeuds_Ecart.NŒUDS)  "
Sql = Sql & "AND (T_Noeuds.Id_IndiceProjet = T_Noeuds_Ecart.Id_IndiceProjet) "
Sql = Sql & "WHERE T_Noeuds_Ecart.Id_IndiceProjet=" & IdIndiceProjet & "  "
Debug.Print Sql

Set RsModifier = Con.OpenRecordSet(Sql)
L = 0: C = 0
If MajEcartExcel(IdIndiceProjet, MyWorkbook, L, C, RsSuprimer, RsAjouter, RsModifier, "Noeuds_Ecart") = True Then boolSave = True


Sql = "SELECT Ligne_Tableau_fils_Ecart.ACTIVER,Ligne_Tableau_fils_Ecart.LIAI, Ligne_Tableau_fils_Ecart.DESIGNATION, Ligne_Tableau_fils_Ecart.FIL, "
Sql = Sql & "Ligne_Tableau_fils_Ecart.SECT, Ligne_Tableau_fils_Ecart.TEINT,  "
Sql = Sql & "Ligne_Tableau_fils_Ecart.TEINT2, Ligne_Tableau_fils_Ecart.ISO,  "
Sql = Sql & "Ligne_Tableau_fils_Ecart.LONG, Ligne_Tableau_fils_Ecart.[LONG CP],  "
Sql = Sql & "Ligne_Tableau_fils_Ecart.COUPE, Ligne_Tableau_fils_Ecart.POS, Ligne_Tableau_fils_Ecart.[POS-OUT],  "
Sql = Sql & "Ligne_Tableau_fils_Ecart.FA, Ligne_Tableau_fils_Ecart.APP, Ligne_Tableau_fils_Ecart.VOI,  "
Sql = Sql & "Ligne_Tableau_fils_Ecart.POS2, Ligne_Tableau_fils_Ecart.[POS-OUT2], Ligne_Tableau_fils_Ecart.FA2,  "
Sql = Sql & "Ligne_Tableau_fils_Ecart.APP2, Ligne_Tableau_fils_Ecart.VOI2, Ligne_Tableau_fils_Ecart.PRECO, Ligne_Tableau_fils_Ecart.OPTION "
Sql = Sql & "FROM Ligne_Tableau_fils_Ecart LEFT JOIN Ligne_Tableau_fils   "
Sql = Sql & "ON (Ligne_Tableau_fils_Ecart.FIL = Ligne_Tableau_fils.FIL)   "
Sql = Sql & "AND (Ligne_Tableau_fils_Ecart.Id_IndiceProjet = Ligne_Tableau_fils.Id_IndiceProjet)  "
Sql = Sql & "WHERE Ligne_Tableau_fils.FIL Is Null AND Ligne_Tableau_fils_Ecart.Id_IndiceProjet=" & IdIndiceProjet & ";  "
Debug.Print Sql


Set RsSuprimer = Con.OpenRecordSet(Sql)

Sql = "SELECT Ligne_Tableau_fils_Ecart.ACTIVER,Ligne_Tableau_fils_Ecart.LIAI, Ligne_Tableau_fils_Ecart.DESIGNATION, Ligne_Tableau_fils_Ecart.FIL,  "
Sql = Sql & "Ligne_Tableau_fils_Ecart.SECT, Ligne_Tableau_fils_Ecart.TEINT, Ligne_Tableau_fils_Ecart.TEINT2,  "
Sql = Sql & "Ligne_Tableau_fils_Ecart.ISO, Ligne_Tableau_fils_Ecart.LONG, Ligne_Tableau_fils_Ecart.[LONG CP],  "
Sql = Sql & "Ligne_Tableau_fils_Ecart.COUPE, Ligne_Tableau_fils_Ecart.POS, Ligne_Tableau_fils_Ecart.[POS-OUT],  "
Sql = Sql & "Ligne_Tableau_fils_Ecart.FA, Ligne_Tableau_fils_Ecart.APP, Ligne_Tableau_fils_Ecart.VOI,  "
Sql = Sql & "Ligne_Tableau_fils_Ecart.POS2, Ligne_Tableau_fils_Ecart.[POS-OUT2], Ligne_Tableau_fils_Ecart.FA2,  "
Sql = Sql & "Ligne_Tableau_fils_Ecart.APP2, Ligne_Tableau_fils_Ecart.VOI2, Ligne_Tableau_fils_Ecart.PRECO, Ligne_Tableau_fils_Ecart.OPTION "
Sql = Sql & "FROM Ligne_Tableau_fils LEFT JOIN Ligne_Tableau_fils_Ecart  "
Sql = Sql & "ON (Ligne_Tableau_fils.FIL = Ligne_Tableau_fils_Ecart.FIL)  "
Sql = Sql & "AND (Ligne_Tableau_fils.Id_IndiceProjet = Ligne_Tableau_fils_Ecart.Id_IndiceProjet) "
Sql = Sql & "WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & IdIndiceProjet & "  "
Sql = Sql & "AND Ligne_Tableau_fils_Ecart.LIAI Is Null;"
Debug.Print Sql
Set RsAjouter = Con.OpenRecordSet(Sql)

Sql = "SELECT '' AS [Avant/Après],Ligne_Tableau_fils_Ecart.ACTIVER, Ligne_Tableau_fils_Ecart.LIAI, Ligne_Tableau_fils_Ecart.DESIGNATION,  "
Sql = Sql & "Ligne_Tableau_fils_Ecart.FIL, Ligne_Tableau_fils_Ecart.SECT, Ligne_Tableau_fils_Ecart.TEINT,  "
Sql = Sql & "Ligne_Tableau_fils_Ecart.TEINT2, Ligne_Tableau_fils_Ecart.ISO, Ligne_Tableau_fils_Ecart.LONG,  "
Sql = Sql & "Ligne_Tableau_fils_Ecart.[LONG CP], Ligne_Tableau_fils_Ecart.COUPE, Ligne_Tableau_fils_Ecart.POS,  "
Sql = Sql & "Ligne_Tableau_fils_Ecart.[POS-OUT], Ligne_Tableau_fils_Ecart.FA, Ligne_Tableau_fils_Ecart.APP,  "
Sql = Sql & "Ligne_Tableau_fils_Ecart.VOI, Ligne_Tableau_fils_Ecart.POS2, Ligne_Tableau_fils_Ecart.[POS-OUT2],  "
Sql = Sql & "Ligne_Tableau_fils_Ecart.FA2, Ligne_Tableau_fils_Ecart.APP2, Ligne_Tableau_fils_Ecart.VOI2,  "
Sql = Sql & "Ligne_Tableau_fils_Ecart.PRECO, Ligne_Tableau_fils_Ecart.OPTION, '' AS Expr1,Ligne_Tableau_fils.ACTIVER, Ligne_Tableau_fils.LIAI,  "
Sql = Sql & "Ligne_Tableau_fils.DESIGNATION, Ligne_Tableau_fils.FIL, Ligne_Tableau_fils.SECT,  "
Sql = Sql & "Ligne_Tableau_fils.TEINT, Ligne_Tableau_fils.TEINT2, Ligne_Tableau_fils.ISO,  "
Sql = Sql & "Ligne_Tableau_fils.LONG, Ligne_Tableau_fils.[LONG CP], Ligne_Tableau_fils.COUPE, Ligne_Tableau_fils.POS,  "
Sql = Sql & "Ligne_Tableau_fils.[POS-OUT], Ligne_Tableau_fils.FA, Ligne_Tableau_fils.APP, Ligne_Tableau_fils.VOI,  "
Sql = Sql & "Ligne_Tableau_fils.POS2, Ligne_Tableau_fils.[POS-OUT2], Ligne_Tableau_fils.FA2,  "
Sql = Sql & "Ligne_Tableau_fils.APP2, Ligne_Tableau_fils.VOI2, Ligne_Tableau_fils.PRECO, Ligne_Tableau_fils.OPTION "
Sql = Sql & "FROM Ligne_Tableau_fils INNER JOIN Ligne_Tableau_fils_Ecart  "
Sql = Sql & "ON (Ligne_Tableau_fils.Id_IndiceProjet = Ligne_Tableau_fils_Ecart.Id_IndiceProjet)  "
Sql = Sql & "AND (Ligne_Tableau_fils.FIL = Ligne_Tableau_fils_Ecart.FIL) "
Sql = Sql & "WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & IdIndiceProjet & ";  "
Debug.Print Sql

Set RsModifier = Con.OpenRecordSet(Sql)

L = 0: C = 0
If MajEcartExcel(IdIndiceProjet, MyWorkbook, L, C, RsSuprimer, RsAjouter, RsModifier, "Tableau_fils_Ecart") = True Then boolSave = True

Sql = "SELECT Connecteurs_Ecart.ACTIVER,Connecteurs_Ecart.CONNECTEUR, Connecteurs_Ecart.[O/N], Connecteurs_Ecart.DESIGNATION,  "
Sql = Sql & "Connecteurs_Ecart.CODE_APP, Connecteurs_Ecart.N°, Connecteurs_Ecart.POS, Connecteurs_Ecart.[POS-OUT],  "
Sql = Sql & "Connecteurs_Ecart.PRECO1, Connecteurs_Ecart.PRECO2, Connecteurs_Ecart.[100%], Connecteurs_Ecart.OPTION "
Sql = Sql & "FROM Connecteurs_Ecart LEFT JOIN Connecteurs  "
Sql = Sql & "ON (Connecteurs_Ecart.Id_IndiceProjet = Connecteurs.Id_IndiceProjet)  "
Sql = Sql & "AND (Connecteurs_Ecart.N° = Connecteurs.N°) "
Sql = Sql & "WHERE Connecteurs_Ecart.Id_IndiceProjet=" & IdIndiceProjet & "  "
Sql = Sql & "AND Connecteurs.CODE_APP Is Null;"
Debug.Print Sql

Set RsSuprimer = Con.OpenRecordSet(Sql)

Sql = "SELECT Connecteurs.ACTIVER,Connecteurs.CONNECTEUR, Connecteurs.[O/N], Connecteurs.DESIGNATION, Connecteurs.CODE_APP,  "
Sql = Sql & "Connecteurs.N°, Connecteurs.POS, Connecteurs.[POS-OUT], Connecteurs.PRECO1, Connecteurs.PRECO2,  "
Sql = Sql & "Connecteurs.[100%], Connecteurs.OPTION "
Sql = Sql & "FROM Connecteurs LEFT JOIN Connecteurs_Ecart  "
Sql = Sql & "ON (Connecteurs.N° = Connecteurs_Ecart.N°)  "
Sql = Sql & "AND (Connecteurs.Id_IndiceProjet = Connecteurs_Ecart.Id_IndiceProjet) "
Sql = Sql & "WHERE Connecteurs.Id_IndiceProjet=" & IdIndiceProjet & "  "
Sql = Sql & "AND Connecteurs_Ecart.CODE_APP Is Null;"
Debug.Print Sql

Set RsAjouter = Con.OpenRecordSet(Sql)

Sql = "SELECT '' AS [Avant/Après], Connecteurs_Ecart.ACTIVER,Connecteurs_Ecart.CONNECTEUR, Connecteurs_Ecart.[O/N],  "
Sql = Sql & "Connecteurs_Ecart.DESIGNATION, Connecteurs_Ecart.CODE_APP, Connecteurs_Ecart.N°,  "
Sql = Sql & "Connecteurs_Ecart.POS, Connecteurs_Ecart.[POS-OUT], Connecteurs_Ecart.PRECO1,  "
Sql = Sql & "Connecteurs_Ecart.PRECO2, Connecteurs_Ecart.[100%], Connecteurs_Ecart.OPTION, Connecteurs_Ecart.[Pylone], Connecteurs_Ecart.[Colonne], Connecteurs_Ecart.[Ligne], '' AS Expr1,  "
Sql = Sql & "Connecteurs.ACTIVER,Connecteurs.CONNECTEUR, Connecteurs.[O/N], Connecteurs.DESIGNATION, Connecteurs.CODE_APP,  "
Sql = Sql & "Connecteurs.N°, Connecteurs.POS, Connecteurs.[POS-OUT], Connecteurs.PRECO1, Connecteurs.PRECO2,  "
Sql = Sql & "Connecteurs.[100%], Connecteurs.OPTION , Connecteurs.[Pylone], Connecteurs.[Colonne], Connecteurs.[Ligne]"
Sql = Sql & "FROM Connecteurs INNER JOIN Connecteurs_Ecart ON (Connecteurs.N° = Connecteurs_Ecart.N°) AND (Connecteurs.Id_IndiceProjet = Connecteurs_Ecart.Id_IndiceProjet) "
Sql = Sql & "WHERE Connecteurs.Id_IndiceProjet=" & IdIndiceProjet & "  "

Debug.Print Sql

Set RsModifier = Con.OpenRecordSet(Sql)

L = 0: C = 0
If MajEcartExcel(IdIndiceProjet, MyWorkbook, L, C, RsSuprimer, RsAjouter, RsModifier, "Connecteurs_Ecart") = True Then boolSave = True

Sql = "SELECT T_Critères_Ecart.ACTIVER,T_Critères_Ecart.CODE_CRITERE, T_Critères_Ecart.CRITERES "
Sql = Sql & "FROM T_Critères_Ecart LEFT JOIN T_Critères ON (T_Critères_Ecart.Id_IndiceProjet = T_Critères.Id_IndiceProjet) "
Sql = Sql & "AND (T_Critères_Ecart.CODE_CRITERE = T_Critères.CODE_CRITERE) "
Sql = Sql & "WHERE T_Critères_Ecart.Id_IndiceProjet=" & IdIndiceProjet & "  "
Sql = Sql & "AND T_Critères.CODE_CRITERE Is Null;"
Debug.Print Sql

Set RsSuprimer = Con.OpenRecordSet(Sql)

Sql = "SELECT T_Critères.ACTIVER,T_Critères.CODE_CRITERE, T_Critères.CRITERES "
Sql = Sql & "FROM T_Critères LEFT JOIN T_Critères_Ecart ON (T_Critères.Id_IndiceProjet =  "
Sql = Sql & "T_Critères_Ecart.Id_IndiceProjet) AND (T_Critères.CODE_CRITERE = T_Critères_Ecart.CODE_CRITERE)"
Sql = Sql & "WHERE T_Critères.Id_IndiceProjet=" & IdIndiceProjet & "  "
Sql = Sql & "AND T_Critères_Ecart.CODE_CRITERE Is Null;"
Debug.Print Sql

Set RsAjouter = Con.OpenRecordSet(Sql)

Sql = "SELECT '' AS [Avant/Après], T_Critères_Ecart.ACTIVER,T_Critères_Ecart.CODE_CRITERE,  "
Sql = Sql & "T_Critères_Ecart.CRITERES, '' AS Expr1, T_Critères.ACTIVER,T_Critères.CODE_CRITERE, T_Critères.CRITERES "
Sql = Sql & "FROM T_Critères INNER JOIN T_Critères_Ecart ON (T_Critères.Id_IndiceProjet =  "
Sql = Sql & "T_Critères_Ecart.Id_IndiceProjet) AND (T_Critères.CODE_CRITERE = T_Critères_Ecart.CODE_CRITERE) "
Sql = Sql & "WHERE T_Critères.Id_IndiceProjet=" & IdIndiceProjet & "  "
Sql = Sql & "ORDER BY T_Critères.CODE_CRITERE;"
Debug.Print Sql

Set RsModifier = Con.OpenRecordSet(Sql)

L = 0: C = 0
If MajEcartExcel(IdIndiceProjet, MyWorkbook, L, C, RsSuprimer, RsAjouter, RsModifier, "Critères_Ecart") = True Then boolSave = True
If boolSave = True Then
Sql = "SELECT RqCartouche.* "
Sql = Sql & "FROM RqCartouche "
 Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"
Set RsModifier = Con.OpenRecordSet(Sql)
If RsModifier.EOF = False Then
        PathPl = PathArchive(PathArchiveAutocad, "" & RsModifier!Client, "" & RsModifier!CleAc, "" & RsModifier!Pieces, "Li", RsModifier.Fields("Li"), IdIndiceProjet, RsModifier.Fields("PI_Indice"), RsModifier.Fields("LI_Indice"), RsModifier!Version, True)
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
        MyWorkbook.SaveAs PathPl & "_Ecart_" & MyFormatDate
        
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
MyEcel.Quit
End Sub
Function sqlRange(Myrange As EXCEL.Range, Row, FROM)
Dim Sql1 As String
Dim Sql2 As String
Dim Sql1Val As String
Dim Sql2Val As String
Dim Sql3Val As String
Sql1 = "INSERT INTO " & FROM & " (Job,"
Sql1Val = ""
Sql2Val = NmJob & ","
Sql3Val = ""
'Myrange.Application.Visible = True
For i = 1 To Myrange.Columns.Count
     If FROM = "Xls_Composants" And i > 4 Then Exit For
    If Trim("" & Myrange(1, i)) = "" Then Exit For
DoEvents
    Sql1Val = Sql1Val & "[" & Myrange(1, i) & "],"
   
    If Trim("" & Myrange(Row, i)) = "" Then
        If Myrange(1, i) = "O/N" Then
             Sql2Val = Sql2Val & "0,"
        Else
            Sql2Val = Sql2Val & "NULL,"
        End If
    Else
        If Myrange(1, i) = "O/N" Then
            If Left(UCase(Trim(Myrange(Row, i))), 1) = "N" Then Myrange(Row, i) = 0
            If Myrange(Row, i) <> 0 Then Myrange(Row, i) = 1
            
            Sql2Val = Sql2Val & "" & CInt(Trim(Myrange(Row, i))) & ","
        Else
            Sql2Val = Sql2Val & "'" & MyReplace(Trim(Myrange(Row, i))) & "',"
        End If
    End If
   
Next i
If FROM = "Xls_Composants" Then
Sql1Val = Sql1Val & "[Path],"
Sql3Val = "NULL,"
    For i = i To Myrange.Columns.Count
        If Val((Trim("" & Myrange(Row, i).Value))) = 1 Then
            Sql3Val = "'" & MyReplace(Trim(Myrange(1, i))) & "',"
            Exit For
        End If
    Next i
End If

Sql1Val = Left(Sql1Val, Len(Sql1Val) - 1)
Sql2Val = Sql2Val & Sql3Val
Sql2Val = Left(Sql2Val, Len(Sql2Val) - 1)
sqlRange = Sql1 & Sql1Val & ") Values(" & Sql2Val & ");"
End Function

Public Sub CeerFichierXls(Xls As String)
Dim Sql As String
Dim Rs As Recordset
Dim IdProjet As Long
Dim IdIndice As Long
DoEvents

Dim MyEcel As New EXCEL.Application
Dim MyClasseur As EXCEL.Workbook
Dim MySheet As EXCEL.Worksheet
Dim Myrange As EXCEL.Range
Set MyClasseur = MyEcel.Workbooks.Open(TableauPath.Item("PathModelXls") & "Ligne_Tableau_fils.xlt")
MyEcel.DisplayAlerts = False
MyClasseur.SaveAs Xls
MyEcel.DisplayAlerts = True
'MyEcel.Visible = True
End Sub
Public Function MajEcartConnecteur(IdIndiceProjet As Long, MyWorkbook As EXCEL.Workbook, L As Long, C As Long, RsSuprimer As Recordset, RsAjouter As Recordset, RsModifier As Recordset) As Boolean
Dim Sql As String
Dim Rs As Recordset

Dim MySheet As EXCEL.Worksheet
Dim Myrange As Range
Sql = "SELECT Connecteurs_Ecart.* "
Sql = Sql & "FROM Connecteurs_Ecart LEFT JOIN Connecteurs  "
Sql = Sql & "ON (Connecteurs_Ecart.Id_IndiceProjet = Connecteurs.Id_IndiceProjet)  "
Sql = Sql & "AND (Connecteurs_Ecart.N° = Connecteurs.N°) "
Sql = Sql & "WHERE Connecteurs_Ecart.Id_IndiceProjet=" & IdIndiceProjet & "  "
Sql = Sql & "AND Connecteurs.CODE_APP Is Null;"
 
 

Set Rs = Con.OpenRecordSet(Sql)
Set MySheet = IsertSheet(MyWorkbook, "Connecteur_Ecart")
Set Myrange = MySheet.Cells(1, 1).CurrentRegion
L = Myrange.Rows.Count
If L > 1 Then L = L + 1
If Rs.EOF = False Then
MajEcartConnecteur = True
Myrange(Myrange.Rows.Count, 1) = "Enregistrement Suprimer"
FormatExcelPlage MySheet.Range(Myrange(L, 1).Address & ":" & Myrange(L, Rs.Fields.Count).Address), 40, True, True, xlCenter, xlCenter


Set Myrange = MySheet.Cells(1, 1).CurrentRegion
L = Myrange.Rows.Count + 1
For i = 0 To Rs.Fields.Count - 1
    Myrange(L, i + 1) = Rs(i).Name
Next
FormatExcelPlage MySheet.Range(Myrange(L, 1).Address & ":" & Myrange(L, Rs.Fields.Count).Address), 15, False, True, xlCenter, xlCenter

While Rs.EOF = False
L = L + 1
   For i = 0 To Rs.Fields.Count - 1
    Myrange(L, i + 1) = "" & Rs(i).Value
    Next
    FormatExcelPlage MySheet.Range(Myrange(L, 1).Address & ":" & Myrange(L, Rs.Fields.Count).Address), 2, False, True, xlCenter, xlCenter

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
Set Myrange = MySheet.Cells(1, 1).CurrentRegion
L = Myrange.Rows.Count
If L > 1 Then L = L + 1
If Rs.EOF = False Then
MajEcartConnecteur = True
Myrange(L, 1) = "Enregistrement Ajouter"
FormatExcelPlage MySheet.Range(Myrange(L, 1).Address & ":" & Myrange(L, Rs.Fields.Count).Address), 40, True, True, xlCenter, xlCenter


Set Myrange = MySheet.Cells(1, 1).CurrentRegion
L = Myrange.Rows.Count + 1

For i = 0 To Rs.Fields.Count - 1
    Myrange(L, i + 1) = Rs(i).Name
Next
FormatExcelPlage MySheet.Range(Myrange(L, 1).Address & ":" & Myrange(L, Rs.Fields.Count).Address), 15, False, True, xlCenter, xlCenter

While Rs.EOF = False
L = L + 1
   For i = 0 To Rs.Fields.Count - 1
    Myrange(L, i + 1) = "" & Rs(i).Value
    Next
    FormatExcelPlage MySheet.Range(Myrange(L, 1).Address & ":" & Myrange(L, Rs.Fields.Count).Address), 2, False, True, xlCenter, xlCenter

    Rs.MoveNext
Wend



End If
Set Rs = Con.CloseRecordSet(Rs)
Sql = "SELECT Connecteurs_Ecart.*, Connecteurs.* "
Sql = Sql & "FROM Connecteurs INNER JOIN Connecteurs_Ecart ON (Connecteurs.N° = Connecteurs_Ecart.N°) AND (Connecteurs.Id_IndiceProjet = Connecteurs_Ecart.Id_IndiceProjet) "
Sql = Sql & "WHERE Connecteurs.Id_IndiceProjet=" & IdIndiceProjet & "  "


Set Rs = Con.OpenRecordSet(Sql)
Set MySheet = IsertSheet(MyWorkbook, "Connecteur_Ecart")
Set Myrange = MySheet.Cells(1, 1).CurrentRegion
L = Myrange.Rows.Count
If L > 1 Then L = L + 1
If Rs.EOF = False Then
Myrange(L, 1) = "Enregistrement Modifier"
FormatExcelPlage MySheet.Range(Myrange(L, 1).Address & ":" & Myrange(L, (Rs.Fields.Count / 2)).Address), 40, True, True, xlCenter, xlCenter


Set Myrange = MySheet.Cells(1, 1).CurrentRegion
L = Myrange.Rows.Count + 1
For i = 0 To (Rs.Fields.Count / 2) - 1
    Myrange(L, i + 1) = Rs(i).Name
Next
FormatExcelPlage MySheet.Range(Myrange(L, 1).Address & ":" & Myrange(L, Rs.Fields.Count / 2).Address), 15, False, True, xlCenter, xlCenter

While Rs.EOF = False
L = L + 1
modifire = False
   For i = 0 To (Rs.Fields.Count - 2) / 2
    If "" & Rs(i).Value <> "" & Rs(13 + i).Value Then
        If i > 0 Then
        MajEcartConnecteur = True
        modifire = True
        Exit For
        End If
    End If
   
    Next
    If modifire = True Then
   For i = 0 To (Rs.Fields.Count - 2) / 2
    Myrange(L, i + 1) = "" & Rs(i).Value & Chr(10) & "" & Rs(13 + i).Value
    If "" & Rs(i).Value <> "" & Rs(13 + i).Value And i > 0 Then
        Myrange(L, i + 1).Font.ColorIndex = 3
       
    End If
    
    Next
          FormatExcelPlage MySheet.Range(Myrange(L, 1).Address & ":" & Myrange(L, Rs.Fields.Count / 2).Address), 2, False, True, xlCenter, xlCenter
    Else
        L = L - 1
    End If
  

    Rs.MoveNext
Wend
End If

Set Rs = Con.CloseRecordSet(Rs)
Set Myrange = Myrange(1, 1).CurrentRegion
If Myrange.Rows.Count = 1 Then Exit Function

Sql = "SELECT  T_indiceProjet.Ensemble, T_indiceProjet.Equipement,[Li] & '_' & [LI_Indice] AS Liste, "
Sql = Sql & "[PI] & '_' & [PI_Indice] AS Piece, T_indiceProjet.Client, T_indiceProjet.CleAc "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"
Dim RsEntetePage As Recordset
Set RsEntetePage = Con.OpenRecordSet(Sql)

MiseEnPage MySheet, Myrange, "Affaire : " & RsEntetePage!CleAc & vbCrLf & _
    "Câblage : " & Replace("" & RsEntetePage!Ensemble, vbCrLf, " ") & vbCrLf & _
     "Pièce : " & RsEntetePage!Piece, vbCrLf & "Equipement : " & Replace("" & RsEntetePage!Equipement, vbCrLf, " ") & vbCrLf & _
     "Liste : " & RsEntetePage!Liste _
    , "ENCELADE" & vbCrLf & "Client : " & RsEntetePage!Client & vbCrLf & Format(Date, "dd-mmm-yyyy"), "", "&P/&N", "", 80, "", False, 2, False, True, True

    MaJEncadreXls Myrange, xlThin, xlThin, xlHairline, xlHairline


Set RsEntetePage = Con.CloseRecordSet(RsEntetePage)
End Function
Public Function MajEcartCritaire(IdIndiceProjet As Long, MyWorkbook As EXCEL.Workbook, L As Long, C As Long, RsSuprimer As Recordset, RsAjouter As Recordset, RsModifier As Recordset) As Boolean
Dim Sql As String
Dim Rs As Recordset

Dim MySheet As EXCEL.Worksheet
Dim Myrange As Range
Sql = "SELECT T_Critères_Ecart.*  "
Sql = Sql & "FROM T_Critères_Ecart LEFT JOIN T_Critères ON (T_Critères_Ecart.Id_IndiceProjet = T_Critères.Id_IndiceProjet) "
Sql = Sql & "AND (T_Critères_Ecart.CODE_CRITERE = T_Critères.CODE_CRITERE) "
Sql = Sql & "WHERE T_Critères_Ecart.Id_IndiceProjet=" & IdIndiceProjet & "  "
Sql = Sql & "AND T_Critères.CODE_CRITERE Is Null;"

Set Rs = Con.OpenRecordSet(Sql)
Set MySheet = IsertSheet(MyWorkbook, "Critères_Ecart")
Set Myrange = MySheet.Cells(1, 1).CurrentRegion
L = Myrange.Rows.Count
If L > 1 Then L = L + 1
If Rs.EOF = False Then
MajEcartCritaire = True
Myrange(Myrange.Rows.Count, 1) = "Enregistrement Suprimer"
FormatExcelPlage MySheet.Range(Myrange(L, 1).Address & ":" & Myrange(L, Rs.Fields.Count).Address), 40, True, True, xlCenter, xlCenter


Set Myrange = MySheet.Cells(1, 1).CurrentRegion
L = Myrange.Rows.Count + 1
For i = 0 To Rs.Fields.Count - 1
    Myrange(L, i + 1) = Rs(i).Name
Next
FormatExcelPlage MySheet.Range(Myrange(L, 1).Address & ":" & Myrange(L, Rs.Fields.Count).Address), 15, False, True, xlCenter, xlCenter

While Rs.EOF = False
L = L + 1
   For i = 0 To Rs.Fields.Count - 1
    Myrange(L, i + 1) = "" & Rs(i).Value
    Next
    FormatExcelPlage MySheet.Range(Myrange(L, 1).Address & ":" & Myrange(L, Rs.Fields.Count).Address), 2, False, True, xlCenter, xlCenter

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
Set Myrange = MySheet.Cells(1, 1).CurrentRegion
L = Myrange.Rows.Count
If L > 1 Then L = L + 1
If Rs.EOF = False Then
MajEcartCritaire = True
Myrange(L, 1) = "Enregistrement Ajouter"
FormatExcelPlage MySheet.Range(Myrange(L, 1).Address & ":" & Myrange(L, Rs.Fields.Count).Address), 40, True, True, xlCenter, xlCenter


Set Myrange = MySheet.Cells(1, 1).CurrentRegion
L = Myrange.Rows.Count + 1

For i = 0 To Rs.Fields.Count - 1
    Myrange(L, i + 1) = Rs(i).Name
Next
FormatExcelPlage MySheet.Range(Myrange(L, 1).Address & ":" & Myrange(L, Rs.Fields.Count).Address), 15, False, True, xlCenter, xlCenter

While Rs.EOF = False
L = L + 1
   For i = 0 To Rs.Fields.Count - 1
    Myrange(L, i + 1) = "" & Rs(i).Value
    Next
    FormatExcelPlage MySheet.Range(Myrange(L, 1).Address & ":" & Myrange(L, Rs.Fields.Count).Address), 2, False, True, xlCenter, xlCenter

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
Set Myrange = MySheet.Cells(1, 1).CurrentRegion
L = Myrange.Rows.Count
If L > 1 Then L = L + 1
If Rs.EOF = False Then
Myrange(L, 1) = "Enregistrement Modifier"
FormatExcelPlage MySheet.Range(Myrange(L, 1).Address & ":" & Myrange(L, (Rs.Fields.Count / 2)).Address), 40, True, True, xlCenter, xlCenter


Set Myrange = MySheet.Cells(1, 1).CurrentRegion
L = Myrange.Rows.Count + 1
For i = 0 To (Rs.Fields.Count / 2) - 1
    Myrange(L, i + 1) = Rs(i).Name
Next
FormatExcelPlage MySheet.Range(Myrange(L, 1).Address & ":" & Myrange(L, Rs.Fields.Count / 2).Address), 15, False, True, xlCenter, xlCenter

While Rs.EOF = False
L = L + 1
modifire = False
   For i = 0 To (Rs.Fields.Count - 2) / 2
    If "" & Rs(i).Value <> "" & Rs(4 + i).Value Then
        If i > 0 Then
        modifire = True
        MajEcartCritaire = True
        Exit For
        End If
    End If
   
    Next
    If modifire = True Then
   For i = 0 To (Rs.Fields.Count - 2) / 2
    Myrange(L, i + 1) = "" & Rs(i).Value & Chr(10) & "" & Rs(4 + i).Value
    If "" & Rs(i).Value <> "" & Rs(4 + i).Value And i > 0 Then
        Myrange(L, i + 1).Font.ColorIndex = 3
       
    End If
    
    Next
          FormatExcelPlage MySheet.Range(Myrange(L, 1).Address & ":" & Myrange(L, Rs.Fields.Count / 2).Address), 2, False, True, xlCenter, xlCenter
    Else
        L = L - 1
    End If
  

    Rs.MoveNext
Wend
End If
Set Rs = Con.CloseRecordSet(Rs)
Set Myrange = Myrange(1, 1).CurrentRegion
If Myrange.Rows.Count = 1 Then Exit Function
Sql = "SELECT  T_indiceProjet.Ensemble, T_indiceProjet.Equipement,[Li] & '_' & [LI_Indice] AS Liste, "
Sql = Sql & "[PI] & '_' & [PI_Indice] AS Piece, T_indiceProjet.Client, T_indiceProjet.CleAc "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"
Dim RsEntetePage As Recordset
Set RsEntetePage = Con.OpenRecordSet(Sql)
Set Myrange = Myrange(1, 1).CurrentRegion

MiseEnPage MySheet, Myrange, "Affaire : " & RsEntetePage!CleAc & vbCrLf & _
    "Câblage : " & Replace("" & RsEntetePage!Ensemble, vbCrLf, " ") & vbCrLf & _
     "Pièce : " & RsEntetePage!Piece, vbCrLf & "Equipement : " & Replace("" & RsEntetePage!Equipement, vbCrLf, " ") & vbCrLf & _
     "Liste : " & RsEntetePage!Liste _
    , "ENCELADE" & vbCrLf & "Client : " & RsEntetePage!Client & vbCrLf & Format(Date, "dd-mmm-yyyy"), "", "&P/&N", "", 80, "", False, 2, False, True, True
    
     MaJEncadreXls Myrange, xlThin, xlThin, xlHairline, xlHairline

Set RsEntetePage = Con.CloseRecordSet(RsEntetePage)
End Function

Public Function MajEcartFils(IdIndiceProjet As Long, MyWorkbook As EXCEL.Workbook, L As Long, C As Long, RsSuprimer As Recordset, RsAjouter As Recordset, RsModifier As Recordset) As Boolean
Dim Sql As String
Dim Rs As Recordset

Dim MySheet As EXCEL.Worksheet
Dim Myrange As Range
Sql = "SELECT Ligne_Tableau_fils_Ecart.*  "
Sql = Sql & "FROM Ligne_Tableau_fils_Ecart LEFT JOIN Ligne_Tableau_fils   "
Sql = Sql & "ON (Ligne_Tableau_fils_Ecart.FIL = Ligne_Tableau_fils.FIL)   "
Sql = Sql & "AND (Ligne_Tableau_fils_Ecart.Id_IndiceProjet = Ligne_Tableau_fils.Id_IndiceProjet)  "
Sql = Sql & "WHERE Ligne_Tableau_fils.FIL Is Null AND Ligne_Tableau_fils_Ecart.Id_IndiceProjet=" & IdIndiceProjet & ";  "

 
 

Set Rs = Con.OpenRecordSet(Sql)
Set MySheet = IsertSheet(MyWorkbook, "Tableau_Fils_Ecart")
Set Myrange = MySheet.Cells(1, 1).CurrentRegion
L = Myrange.Rows.Count
If L > 1 Then L = L + 1
If Rs.EOF = False Then
MajEcartFils = True
Myrange(Myrange.Rows.Count, 1) = "Enregistrement Suprimer"
FormatExcelPlage MySheet.Range(Myrange(L, 1).Address & ":" & Myrange(L, Rs.Fields.Count).Address), 40, True, True, xlCenter, xlCenter


Set Myrange = MySheet.Cells(1, 1).CurrentRegion
L = Myrange.Rows.Count + 1
For i = 0 To Rs.Fields.Count - 1
    Myrange(L, i + 1) = Rs(i).Name
Next
FormatExcelPlage MySheet.Range(Myrange(L, 1).Address & ":" & Myrange(L, Rs.Fields.Count).Address), 15, False, True, xlCenter, xlCenter

While Rs.EOF = False
L = L + 1
   For i = 0 To Rs.Fields.Count - 1
    Myrange(L, i + 1) = "" & Rs(i).Value
    Next
    FormatExcelPlage MySheet.Range(Myrange(L, 1).Address & ":" & Myrange(L, Rs.Fields.Count).Address), 2, False, True, xlCenter, xlCenter

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
Set Myrange = MySheet.Cells(1, 1).CurrentRegion
L = Myrange.Rows.Count
If L > 1 Then L = L + 1
If Rs.EOF = False Then
MajEcartFils = True
Myrange(L, 1) = "Enregistrement Ajouter"
FormatExcelPlage MySheet.Range(Myrange(L, 1).Address & ":" & Myrange(L, Rs.Fields.Count).Address), 40, True, True, xlCenter, xlCenter


Set Myrange = MySheet.Cells(1, 1).CurrentRegion
L = Myrange.Rows.Count + 1

For i = 0 To Rs.Fields.Count - 1
    Myrange(L, i + 1) = Rs(i).Name
Next
FormatExcelPlage MySheet.Range(Myrange(L, 1).Address & ":" & Myrange(L, Rs.Fields.Count).Address), 15, False, True, xlCenter, xlCenter

While Rs.EOF = False
L = L + 1
   For i = 0 To Rs.Fields.Count - 1
    Myrange(L, i + 1) = "" & Rs(i).Value
    Next
    FormatExcelPlage MySheet.Range(Myrange(L, 1).Address & ":" & Myrange(L, Rs.Fields.Count).Address), 2, False, True, xlCenter, xlCenter

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
Set Myrange = MySheet.Cells(1, 1).CurrentRegion
L = Myrange.Rows.Count
If L > 1 Then L = L + 1
If Rs.EOF = False Then
Myrange(L, 1) = "Enregistrement Modifier"
FormatExcelPlage MySheet.Range(Myrange(L, 1).Address & ":" & Myrange(L, (Rs.Fields.Count / 2)).Address), 40, True, True, xlCenter, xlCenter


Set Myrange = MySheet.Cells(1, 1).CurrentRegion
L = Myrange.Rows.Count + 1
For i = 0 To (Rs.Fields.Count / 2) - 1
    Myrange(L, i + 1) = Rs(i).Name
Next
FormatExcelPlage MySheet.Range(Myrange(L, 1).Address & ":" & Myrange(L, Rs.Fields.Count / 2).Address), 15, False, True, xlCenter, xlCenter

While Rs.EOF = False
L = L + 1
modifire = False
   For i = 0 To (Rs.Fields.Count - 2) / 2
    If "" & Rs(i).Value <> "" & Rs(24 + i).Value Then
        If i > 0 Then
        MajEcartFils = True
        modifire = True
        Exit For
        End If
    End If
   
    Next
    If modifire = True Then
   For i = 0 To (Rs.Fields.Count - 2) / 2
    Myrange(L, i + 1) = "" & Rs(i).Value & Chr(10) & "" & Rs(24 + i).Value
    If "" & Rs(i).Value <> "" & Rs(24 + i).Value And i > 0 Then
        Myrange(L, i + 1).Font.ColorIndex = 3
       
    End If
    
    Next
          FormatExcelPlage MySheet.Range(Myrange(L, 1).Address & ":" & Myrange(L, Rs.Fields.Count / 2).Address), 2, False, True, xlCenter, xlCenter
    Else
        L = L - 1
    End If
  

    Rs.MoveNext
Wend
End If
Set Rs = Con.CloseRecordSet(Rs)
Set Myrange = Myrange(1, 1).CurrentRegion
If Myrange.Rows.Count = 1 Then Exit Function

Sql = "SELECT  T_indiceProjet.Ensemble, T_indiceProjet.Equipement,[Li] & '_' & [LI_Indice] AS Liste, "
Sql = Sql & "[PI] & '_' & [PI_Indice] AS Piece, T_indiceProjet.Client, T_indiceProjet.CleAc "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"
Dim RsEntetePage As Recordset
Set RsEntetePage = Con.OpenRecordSet(Sql)

Set Myrange = Myrange(1, 1).CurrentRegion
MiseEnPage MySheet, Myrange, "Affaire : " & RsEntetePage!CleAc & vbCrLf & _
    "Câblage : " & Replace("" & RsEntetePage!Ensemble, vbCrLf, " ") & vbCrLf & _
     "Pièce : " & RsEntetePage!Piece, vbCrLf & "Equipement : " & Replace("" & RsEntetePage!Equipement, vbCrLf, " ") & vbCrLf & _
     "Liste : " & RsEntetePage!Liste _
    , "ENCELADE" & vbCrLf & "Client : " & RsEntetePage!Client & vbCrLf & Format(Date, "dd-mmm-yyyy"), "", "&P/&N", "", 80, "", False, 2, False, True, True

    MaJEncadreXls Myrange, xlThin, xlThin, xlHairline, xlHairline

Set RsEntetePage = Con.CloseRecordSet(RsEntetePage)
End Function
Public Function MajEcartNota(IdIndiceProjet As Long, MyWorkbook As EXCEL.Workbook, L As Long, C As Long, RsSuprimer As Recordset, RsAjouter As Recordset, RsModifier As Recordset) As Boolean
Dim Sql As String
Dim Rs As Recordset

Dim MySheet As EXCEL.Worksheet
Dim Myrange As Range
Sql = "SELECT Nota_Ecart.* "
Sql = Sql & "FROM Nota_Ecart LEFT JOIN Nota ON (Nota_Ecart.NUMNOTA = Nota.NUMNOTA) AND (Nota_Ecart.Id_IndiceProjet = Nota.Id_IndiceProjet) "
Sql = Sql & "WHERE Nota_Ecart.Id_IndiceProjet=" & IdIndiceProjet & " "
Sql = Sql & "AND Nota.NUMNOTA Is Null;"


Set Rs = Con.OpenRecordSet(Sql)
Set MySheet = IsertSheet(MyWorkbook, "Notas_Ecart")
Set Myrange = MySheet.Cells(1, 1).CurrentRegion
L = Myrange.Rows.Count
If L > 1 Then L = L + 1
If Rs.EOF = False Then
MajEcartNota = True
Myrange(Myrange.Rows.Count, 1) = "Enregistrement Suprimer"
FormatExcelPlage MySheet.Range(Myrange(L, 1).Address & ":" & Myrange(L, Rs.Fields.Count).Address), 40, True, True, xlCenter, xlCenter


Set Myrange = MySheet.Cells(1, 1).CurrentRegion
L = Myrange.Rows.Count + 1
For i = 0 To Rs.Fields.Count - 1
    Myrange(L, i + 1) = Rs(i).Name
Next
FormatExcelPlage MySheet.Range(Myrange(L, 1).Address & ":" & Myrange(L, Rs.Fields.Count).Address), 15, False, True, xlCenter, xlCenter

While Rs.EOF = False
L = L + 1
   For i = 0 To Rs.Fields.Count - 1
    Myrange(L, i + 1) = "" & Rs(i).Value
    Next
    FormatExcelPlage MySheet.Range(Myrange(L, 1).Address & ":" & Myrange(L, Rs.Fields.Count).Address), 2, False, True, xlCenter, xlCenter

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
Set Myrange = MySheet.Cells(1, 1).CurrentRegion
L = Myrange.Rows.Count
If L > 1 Then L = L + 1
If Rs.EOF = False Then
MajEcartNota = True
Myrange(L, 1) = "Enregistrement Ajouter"
FormatExcelPlage MySheet.Range(Myrange(L, 1).Address & ":" & Myrange(L, Rs.Fields.Count).Address), 40, True, True, xlCenter, xlCenter


Set Myrange = MySheet.Cells(1, 1).CurrentRegion
L = Myrange.Rows.Count + 1

For i = 0 To Rs.Fields.Count - 1
    Myrange(L, i + 1) = Rs(i).Name
Next
FormatExcelPlage MySheet.Range(Myrange(L, 1).Address & ":" & Myrange(L, Rs.Fields.Count).Address), 15, False, True, xlCenter, xlCenter

While Rs.EOF = False
L = L + 1
   For i = 0 To Rs.Fields.Count - 1
    Myrange(L, i + 1) = "" & Rs(i).Value
    Next
    FormatExcelPlage MySheet.Range(Myrange(L, 1).Address & ":" & Myrange(L, Rs.Fields.Count).Address), 2, False, True, xlCenter, xlCenter

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
Set Myrange = MySheet.Cells(1, 1).CurrentRegion
L = Myrange.Rows.Count
If L > 1 Then L = L + 1
If Rs.EOF = False Then
Myrange(L, 1) = "Enregistrement Modifier"
FormatExcelPlage MySheet.Range(Myrange(L, 1).Address & ":" & Myrange(L, (Rs.Fields.Count / 2)).Address), 40, True, True, xlCenter, xlCenter


Set Myrange = MySheet.Cells(1, 1).CurrentRegion
L = Myrange.Rows.Count + 1
For i = 0 To (Rs.Fields.Count / 2) - 1
    Myrange(L, i + 1) = Rs(i).Name
Next
FormatExcelPlage MySheet.Range(Myrange(L, 1).Address & ":" & Myrange(L, Rs.Fields.Count / 2).Address), 15, False, True, xlCenter, xlCenter

While Rs.EOF = False
L = L + 1
modifire = False
   For i = 0 To (Rs.Fields.Count - 2) / 2
    If "" & Rs(i).Value <> "" & Rs(4 + i).Value Then
        If i > 0 Then
        modifire = True
        MajEcartNota = True
        Exit For
        End If
    End If
   
    Next
    If modifire = True Then
   For i = 0 To (Rs.Fields.Count - 2) / 2
    Myrange(L, i + 1) = "" & Rs(i).Value & Chr(10) & "" & Rs(4 + i).Value
    If "" & Rs(i).Value <> "" & Rs(4 + i).Value And i > 0 Then
        Myrange(L, i + 1).Font.ColorIndex = 3
       
    End If
    
    Next
          FormatExcelPlage MySheet.Range(Myrange(L, 1).Address & ":" & Myrange(L, Rs.Fields.Count / 2).Address), 2, False, True, xlCenter, xlCenter
    Else
        L = L - 1
    End If
  

    Rs.MoveNext
Wend
End If
Set Rs = Con.CloseRecordSet(Rs)
Set Myrange = Myrange(1, 1).CurrentRegion
If Myrange.Rows.Count = 1 Then Exit Function

Sql = "SELECT  T_indiceProjet.Ensemble, T_indiceProjet.Equipement,[Li] & '_' & [LI_Indice] AS Liste, "
Sql = Sql & "[PI] & '_' & [PI_Indice] AS Piece, T_indiceProjet.Client, T_indiceProjet.CleAc "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"
Dim RsEntetePage As Recordset
Set RsEntetePage = Con.OpenRecordSet(Sql)
Set Myrange = Myrange(1, 1).CurrentRegion

MiseEnPage MySheet, Myrange, "Affaire : " & RsEntetePage!CleAc & vbCrLf & _
    "Câblage : " & Replace("" & RsEntetePage!Ensemble, vbCrLf, " ") & vbCrLf & _
     "Pièce : " & RsEntetePage!Piece, vbCrLf & "Equipement : " & Replace("" & RsEntetePage!Equipement, vbCrLf, " ") & vbCrLf & _
     "Liste : " & RsEntetePage!Liste _
    , "ENCELADE" & vbCrLf & "Client : " & RsEntetePage!Client & vbCrLf & Format(Date, "dd-mmm-yyyy"), "", "&P/&N", "", 80, "", False, 2, False, True, True

    MaJEncadreXls Myrange, xlThin, xlThin, xlHairline, xlHairline


Set RsEntetePage = Con.CloseRecordSet(RsEntetePage)
End Function
Public Function MajEcartComposants(IdIndiceProjet As Long, MyWorkbook As EXCEL.Workbook, L As Long, C As Long, RsSuprimer As Recordset, RsAjouter As Recordset, RsModifier As Recordset) As Boolean
Dim Sql As String
Dim Rs As Recordset

Dim MySheet As EXCEL.Worksheet
Dim Myrange As Range
Sql = "SELECT Composants_Ecart.* "
Sql = Sql & "FROM Composants_Ecart LEFT JOIN Composants  "
Sql = Sql & "ON (Composants_Ecart.NUMCOMP = Composants.NUMCOMP)  "
Sql = Sql & "AND (Composants_Ecart.Id_IndiceProjet = Composants.Id_IndiceProjet) "
Sql = Sql & "WHERE Composants.NUMCOMP Is Null "
Sql = Sql & "AND Composants_Ecart.Id_IndiceProjet=" & IdIndiceProjet & " ;"
 

Set Rs = Con.OpenRecordSet(Sql)
Set MySheet = IsertSheet(MyWorkbook, "Composants_Ecart")
Set Myrange = MySheet.Cells(1, 1).CurrentRegion
L = Myrange.Rows.Count
If L > 1 Then L = L + 1
If Rs.EOF = False Then
MajEcartComposants = True
Myrange(Myrange.Rows.Count, 1) = "Enregistrement Suprimer"
FormatExcelPlage MySheet.Range(Myrange(L, 1).Address & ":" & Myrange(L, Rs.Fields.Count).Address), 40, True, True, xlCenter, xlCenter


Set Myrange = MySheet.Cells(1, 1).CurrentRegion
L = Myrange.Rows.Count + 1
For i = 0 To Rs.Fields.Count - 1
    Myrange(L, i + 1) = Rs(i).Name
Next
FormatExcelPlage MySheet.Range(Myrange(L, 1).Address & ":" & Myrange(L, Rs.Fields.Count).Address), 15, False, True, xlCenter, xlCenter

While Rs.EOF = False
L = L + 1
   For i = 0 To Rs.Fields.Count - 1
    Myrange(L, i + 1) = "" & Rs(i).Value
    Next
    FormatExcelPlage MySheet.Range(Myrange(L, 1).Address & ":" & Myrange(L, Rs.Fields.Count).Address), 2, False, True, xlCenter, xlCenter

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
Set Myrange = MySheet.Cells(1, 1).CurrentRegion
L = Myrange.Rows.Count
If L > 1 Then L = L + 1
If Rs.EOF = False Then
MajEcartComposants = True
Myrange(L, 1) = "Enregistrement Ajouter"
FormatExcelPlage MySheet.Range(Myrange(L, 1).Address & ":" & Myrange(L, Rs.Fields.Count).Address), 40, True, True, xlCenter, xlCenter


Set Myrange = MySheet.Cells(1, 1).CurrentRegion
L = Myrange.Rows.Count + 1

For i = 0 To Rs.Fields.Count - 1
    Myrange(L, i + 1) = Rs(i).Name
Next
FormatExcelPlage MySheet.Range(Myrange(L, 1).Address & ":" & Myrange(L, Rs.Fields.Count).Address), 15, False, True, xlCenter, xlCenter

While Rs.EOF = False
L = L + 1
   For i = 0 To Rs.Fields.Count - 1
    Myrange(L, i + 1) = "" & Rs(i).Value
    Next
    FormatExcelPlage MySheet.Range(Myrange(L, 1).Address & ":" & Myrange(L, Rs.Fields.Count).Address), 2, False, True, xlCenter, xlCenter

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
Set Myrange = MySheet.Cells(1, 1).CurrentRegion
L = Myrange.Rows.Count
If L > 1 Then L = L + 1
If Rs.EOF = False Then
Myrange(L, 1) = "Enregistrement Modifier"
FormatExcelPlage MySheet.Range(Myrange(L, 1).Address & ":" & Myrange(L, (Rs.Fields.Count / 2)).Address), 40, True, True, xlCenter, xlCenter


Set Myrange = MySheet.Cells(1, 1).CurrentRegion
L = Myrange.Rows.Count + 1
For i = 0 To (Rs.Fields.Count / 2) - 1
    Myrange(L, i + 1) = Rs(i).Name
Next
FormatExcelPlage MySheet.Range(Myrange(L, 1).Address & ":" & Myrange(L, Rs.Fields.Count / 2).Address), 15, False, True, xlCenter, xlCenter

While Rs.EOF = False
L = L + 1
modifire = False
   For i = 0 To (Rs.Fields.Count - 2) / 2
    If "" & Rs(i).Value <> "" & Rs(6 + i).Value Then
        If i > 0 Then
        MajEcartComposants = True
        modifire = True
        Exit For
        End If
    End If
   
    Next
    If modifire = True Then
   For i = 0 To (Rs.Fields.Count - 2) / 2
    Myrange(L, i + 1) = "" & Rs(i).Value & Chr(10) & "" & Rs(6 + i).Value
    If "" & Rs(i).Value <> "" & Rs(6 + i).Value And i > 0 Then
        Myrange(L, i + 1).Font.ColorIndex = 3
       
    End If
    
    Next
          FormatExcelPlage MySheet.Range(Myrange(L, 1).Address & ":" & Myrange(L, Rs.Fields.Count / 2).Address), 2, False, True, xlCenter, xlCenter
    Else
        L = L - 1
    End If
  

    Rs.MoveNext
Wend
End If
Set Rs = Con.CloseRecordSet(Rs)
Set Myrange = Myrange(1, 1).CurrentRegion
If Myrange.Rows.Count = 1 Then Exit Function

Sql = "SELECT  T_indiceProjet.Ensemble, T_indiceProjet.Equipement,[Li] & '_' & [LI_Indice] AS Liste, "
Sql = Sql & "[PI] & '_' & [PI_Indice] AS Piece, T_indiceProjet.Client, T_indiceProjet.CleAc "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"
Dim RsEntetePage As Recordset
Set RsEntetePage = Con.OpenRecordSet(Sql)
Set Myrange = Myrange(1, 1).CurrentRegion

MiseEnPage MySheet, Myrange, "Affaire : " & RsEntetePage!CleAc & vbCrLf & _
    "Câblage : " & Replace("" & RsEntetePage!Ensemble, vbCrLf, " ") & vbCrLf & _
     "Pièce : " & RsEntetePage!Piece, vbCrLf & "Equipement : " & Replace("" & RsEntetePage!Equipement, vbCrLf, " ") & vbCrLf & _
     "Liste : " & RsEntetePage!Liste _
    , "ENCELADE" & vbCrLf & "Client : " & RsEntetePage!Client & vbCrLf & Format(Date, "dd-mmm-yyyy"), "", "&P/&N", "", 80, "", False, 2, False, True, True

    MaJEncadreXls Myrange, xlThin, xlThin, xlHairline, xlHairline

Set RsEntetePage = Con.CloseRecordSet(RsEntetePage)
End Function


Function MajEcartExcel(IdIndiceProjet As Long, MyWorkbook As EXCEL.Workbook, L As Long, C As Long, RsSuprimer As Recordset, RsAjouter As Recordset, RsModifier As Recordset, SheetName As String, Optional txt = "") As Boolean
Dim Sql As String


Dim MySheet As EXCEL.Worksheet
Dim Myrange As Range
RecherModifier RsModifier
Set MySheet = IsertSheet(MyWorkbook, SheetName)
Set Myrange = MySheet.Cells(1, 1).CurrentRegion
'MyWorkbook.Application.Visible = True
L = Myrange.Rows.Count
If L > 1 Then L = L + 1

If Trim("" & txt) <> "" Then
t_txt = Split(txt, Chr(10))
i2 = 0
For i = LBound(t_txt) To UBound(t_txt)
If Trim("" & t_txt(i)) <> "" Then
Myrange(L + i - i2, 1) = t_txt(i)
FormatExcelPlage MySheet.Range(Myrange(L + i - i2, 1).Address & ":" & Myrange(L + i - i2, RsModifier.Fields.Count / 2).Address), 2, True, False, xlCenter, xlCenter

Else
    i2 = i2 + 1
End If
Next

L = L + 1
End If
If RsSuprimer.EOF = False Then
MajEcartExcel = True
Myrange(Myrange.Rows.Count, 1) = "Enregistrement Suprimer"

Set Myrange = MySheet.Cells(1, 1).CurrentRegion
L = Myrange.Rows.Count + 1
For i = 0 To RsSuprimer.Fields.Count - 1
    Myrange(L, i + 1) = RsSuprimer(i).Name
Next
FormatExcelPlage MySheet.Range(Myrange(L, 1).Address & ":" & Myrange(L, RsSuprimer.Fields.Count).Address), 15, False, True, xlCenter, xlCenter

While RsSuprimer.EOF = False
DoEvents
L = L + 1
   For i = 0 To RsSuprimer.Fields.Count - 1
    Myrange(L, i + 1) = "" & RsSuprimer(i).Value
    Next
    FormatExcelPlage MySheet.Range(Myrange(L, 1).Address & ":" & Myrange(L, RsSuprimer.Fields.Count).Address), 2, False, True, xlCenter, xlCenter

    RsSuprimer.MoveNext
Wend
End If

Set MySheet = IsertSheet(MyWorkbook, SheetName)
Set Myrange = MySheet.Cells(1, 1).CurrentRegion
L = Myrange.Rows.Count
If L > 1 Then
    L = L + 1
Else
    If Trim("" & txt) <> "" Then L = L + 1
End If
If RsAjouter.EOF = False Then
DoEvents
MajEcartExcel = True
Myrange(L, 1) = "Enregistrement Ajouter"
FormatExcelPlage MySheet.Range(Myrange(L, 1).Address & ":" & Myrange(L, RsAjouter.Fields.Count).Address), 40, True, True, xlCenter, xlCenter


Set Myrange = MySheet.Cells(1, 1).CurrentRegion
L = Myrange.Rows.Count + 1

For i = 0 To RsAjouter.Fields.Count - 1
    Myrange(L, i + 1) = RsAjouter(i).Name
Next
FormatExcelPlage MySheet.Range(Myrange(L, 1).Address & ":" & Myrange(L, RsAjouter.Fields.Count).Address), 15, False, True, xlCenter, xlCenter

While RsAjouter.EOF = False
DoEvents
L = L + 1
   For i = 0 To RsAjouter.Fields.Count - 1
    Myrange(L, i + 1) = "" & RsAjouter(i).Value
    Next
    FormatExcelPlage MySheet.Range(Myrange(L, 1).Address & ":" & Myrange(L, RsAjouter.Fields.Count).Address), 2, False, True, xlCenter, xlCenter

    RsAjouter.MoveNext
Wend



End If

Set MySheet = IsertSheet(MyWorkbook, SheetName)
Set Myrange = MySheet.Cells(1, 1).CurrentRegion
L = Myrange.Rows.Count
If L > 1 Then
    L = L + 1
Else
    If Trim("" & txt) <> "" Then L = L + 1
End If

If RsModifier.EOF = False Then
Myrange(L, 1) = "Enregistrement Modifier"
FormatExcelPlage MySheet.Range(Myrange(L, 1).Address & ":" & Myrange(L, (RsModifier.Fields.Count / 2)).Address), 40, True, True, xlCenter, xlCenter


Set Myrange = MySheet.Cells(1, 1).CurrentRegion
L = Myrange.Rows.Count + 1
For i = 0 To (RsModifier.Fields.Count / 2) - 1
    Myrange(L, i + 1) = RsModifier(i).Name
Next
FormatExcelPlage MySheet.Range(Myrange(L, 1).Address & ":" & Myrange(L, RsModifier.Fields.Count / 2).Address), 15, False, True, xlCenter, xlCenter

While RsModifier.EOF = False
L = L + 1
modifire = False
   For i = 0 To (RsModifier.Fields.Count - 2) / 2
    
    If "" & RsModifier(i).Value <> "" & RsModifier(((RsModifier.Fields.Count) / 2) + i).Value Then
       
        modifire = True
        MajEcartExcel = True
        Exit For
        
    End If
   
    Next
    If modifire = True Then
   For i = 0 To (RsModifier.Fields.Count - 2) / 2
   a = RsModifier(i).Name
   aa = RsModifier(((RsModifier.Fields.Count)) / 2 + i).Name
    If UCase(RsModifier.Fields(i).Name) = UCase("Avant/Après") Then
        Myrange(L, i + 1) = "Avant"
        Myrange(L + 1, i + 1) = "Après"
    End If
   If RsModifier(i).Type = adBoolean Then
        If RsModifier(i).Value = True Then
             Myrange(L, i + 1) = "1"  '& RsModifier(((RsModifier.Fields.Count)) / 2 + i).Value
        Else
            Myrange(L, i + 1) = "0"   '& RsModifier(((RsModifier.Fields.Count)) / 2 + i).Value
        End If
        If RsModifier(((RsModifier.Fields.Count)) / 2 + i) = True Then
             Myrange(L + 1, i + 1) = "1" '& RsModifier(((RsModifier.Fields.Count)) / 2 + i).Value
        Else
            Myrange(L + 1, i + 1) = "0" '& RsModifier(((RsModifier.Fields.Count)) / 2 + i).Value
        End If
   Else
   a = RsModifier(i).Name
   aa = RsModifier(((RsModifier.Fields.Count)) / 2 + i).Name
    If UCase(RsModifier.Fields(i).Name) <> UCase("Avant/Après") Then
        Myrange(L, i + 1) = "" & RsModifier(i).Value
        Myrange(L + 1, i + 1) = "" & RsModifier(((RsModifier.Fields.Count)) / 2 + i).Value
    End If
    End If
    If "" & RsModifier(i).Value <> "" & RsModifier(((RsModifier.Fields.Count)) / 2 + i).Value Then
        Myrange(L + 1, i + 1).Font.ColorIndex = 3
       
    End If
    
    Next
          FormatExcelPlage MySheet.Range(Myrange(L, 1).Address & ":" & Myrange(L + 1, RsModifier.Fields.Count / 2).Address), 2, False, True, xlCenter, xlCenter
    Else
        
    End If
  Set Myrange = MySheet.Cells(1, 1).CurrentRegion
    L = Myrange.Rows.Count

    RsModifier.MoveNext
Wend
End If
'Set Rs = Con.CloseRecordSet(Rs)
Set Myrange = Myrange(1, 1).CurrentRegion
If Myrange.Rows.Count = 1 Then Exit Function

Sql = "SELECT  T_indiceProjet.Ensemble, T_indiceProjet.Equipement,[Li] & '_' & [LI_Indice] AS Liste, "
Sql = Sql & "[PI] & '_' & [PI_Indice] AS Piece, T_indiceProjet.Client, T_indiceProjet.CleAc "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"
Dim RsEntetePage As Recordset
Set RsEntetePage = Con.OpenRecordSet(Sql)
Set Myrange = Myrange(1, 1).CurrentRegion

MiseEnPage MySheet, Myrange, "Affaire : " & RsEntetePage!CleAc & vbCrLf & _
    "Câblage : " & Replace("" & RsEntetePage!Ensemble, vbCrLf, " ") & vbCrLf & _
     "Pièce : " & RsEntetePage!Piece, vbCrLf & "Equipement : " & Replace("" & RsEntetePage!Equipement, vbCrLf, " ") & vbCrLf & _
     "Liste : " & RsEntetePage!Liste _
    , "ENCELADE" & vbCrLf & "Client : " & RsEntetePage!Client & vbCrLf & Format(Date, "dd-mmm-yyyy"), "", "&P/&N", "", 80, "", False, 2, False, True, True

    MaJEncadreXls Myrange, xlThin, xlThin, xlHairline, xlHairline

Set RsEntetePage = Con.CloseRecordSet(RsEntetePage)
End Function

Sub RecherModifier(RsModifier As Recordset)

While RsModifier.EOF = False
DoEvents
   For i = 0 To (RsModifier.Fields.Count - 2) / 2
    If "" & RsModifier(i).Value <> "" & RsModifier(((RsModifier.Fields.Count) / 2) + i).Value Then
       
        Exit Sub
        
    End If
   
    Next
    RsModifier.MoveNext
Wend

End Sub


