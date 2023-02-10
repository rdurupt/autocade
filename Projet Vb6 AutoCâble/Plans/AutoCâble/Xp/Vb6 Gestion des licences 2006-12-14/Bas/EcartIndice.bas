Attribute VB_Name = "EcartIndice"
Sub MajEcartIndice(IdIndiceProjet As Long)
 Set TableauPath = funPath
Dim L As Long
Dim C As Long
Dim boolSave As Boolean
Dim Sql As String
Dim Rs As Recordset
Dim RsSuprimer As Recordset
Dim RsAjouter As Recordset
Dim RsModifier As Recordset
Dim MyExcel As New EXCEL.Application
MyExcel.DisplayAlerts = False
Dim MyIndice As Long
Dim MyIndceMoins1 As Long
Dim Fso As New FileSystemObject
boolSave = False
MyExcel.Visible = True
   PathArchiveAutocad = TableauPath.Item("PathArchiveAutocad")
   PathArchiveAutocad = DefinirChemienComplet(TableauPath.Item("PathServer"), PathArchiveAutocad)
'     If Left(PathArchiveAutocad, 2) <> "\\" And Left(PathArchiveAutocad, 1) = "\" Then PathArchiveAutocad = TableauPath.Item("PathServer") & PathArchiveAutocad
'  If Right(PathArchiveAutocad, 2) = "\\" Then PathArchiveAutocad = Mid(PathArchiveAutocad, 1, Len(MyPath) - 1)
'MyExcel.Visible = True
Set MyWorkbook = MyExcel.Workbooks.Add
For I = MyWorkbook.Worksheets.Count To 1 Step -1
    DeletSheet MyWorkbook.Worksheets(I)
Next
Sql = "SELECT TOP 2 T_indiceProjet_1.Id, T_indiceProjet_1.ReffIndice, T_indiceProjet_1.Description,  "
Sql = Sql & "[T_indiceProjet_1].[PI] & '_' & Trim('' & [T_indiceProjet_1].[Pi_Indice]) AS Piece,  "
Sql = Sql & "[T_indiceProjet_1].[Pl] & '_' & Trim('' & [T_indiceProjet_1].[Li_Indice]) AS Plan,  "
Sql = Sql & "[T_indiceProjet_1].[ou] & '_' & Trim('' & [T_indiceProjet_1].[ou_Indice]) AS Outil,  "
Sql = Sql & "[T_indiceProjet_1].[Li] & '_' & Trim('' & [T_indiceProjet_1].[Li_Indice]) AS Liste "
Sql = Sql & "FROM T_indiceProjet INNER JOIN T_indiceProjet AS T_indiceProjet_1  "
Sql = Sql & "ON T_indiceProjet.Id_Pieces = T_indiceProjet_1.Id_Pieces "
Sql = Sql & "Where T_indiceProjet.Id_Pieces=" & IdIndiceProjet & " "
Sql = Sql & "GROUP BY T_indiceProjet_1.Id, T_indiceProjet_1.ReffIndice, T_indiceProjet_1.Description, "
Sql = Sql & "[T_indiceProjet_1].[PI] & '_' & Trim('' & [T_indiceProjet_1].[Pi_Indice]), "
Sql = Sql & "[T_indiceProjet_1].[Pl] & '_' & Trim('' & [T_indiceProjet_1].[Li_Indice]), "
Sql = Sql & "[T_indiceProjet_1].[ou] & '_' & Trim('' & [T_indiceProjet_1].[ou_Indice]), "
Sql = Sql & "[T_indiceProjet_1].[Li] & '_' & Trim('' & [T_indiceProjet_1].[Li_Indice]), T_indiceProjet_1.PI_Indice "
Sql = Sql & "ORDER BY T_indiceProjet_1.Id DESC;"
Set Rs = Con.OpenRecordSet(Sql)
MyIndice = 0
MyIndceMoins1 = 0
If Rs.EOF = False Then
    MyIndice = Rs!Id
    Rs.MoveNext
End If
If Rs.EOF = False Then
    MyIndceMoins1 = Rs!Id
    Rs.MoveNext
End If
Rs.Requery
TableProjet = Rs.GetRows
If UBound(TableProjet, 2) - 1 = -1 Then GoTo Sortie
For I = 0 To UBound(TableProjet, 2) - 1

    Sql = "DROP TABLE Temp_" & NmJob & "_Ecart_Nota;"
    Con.Execute Sql
    Sql = "SELECT  Nota.* INTO Temp_" & NmJob & "_Ecart_Nota "
    Sql = Sql & "FROM Nota "
    Sql = Sql & "WHERE Nota.Id_IndiceProjet=" & MyIndceMoins1 & ";"
    
    Con.Execute Sql
    
    Sql = "SELECT Temp_" & NmJob & "_Ecart_Nota.ACTIVER, Temp_" & NmJob & "_Ecart_Nota.NOTA, Temp_" & NmJob & "_Ecart_Nota.NUMNOTA, Temp_" & NmJob & "_Ecart_Nota.OPTION "
    Sql = Sql & "FROM Nota LEFT JOIN Temp_" & NmJob & "_Ecart_Nota ON Nota.NOTA = Temp_" & NmJob & "_Ecart_Nota.NOTA  "
    Sql = Sql & "WHERE Nota.Id_IndiceProjet=" & MyIndice & " AND Nota.NUMNOTA Is Null;"


    Debug.Print Sql


Set RsSuprimer = Con.OpenRecordSet(Sql)

Sql = "SELECT Nota.ACTIVER,Nota.NOTA, Nota.NUMNOTA ,Nota.[OPTION] "
Sql = Sql & "FROM Temp_" & NmJob & "_Ecart_Nota LEFT JOIN Nota ON Temp_" & NmJob & "_Ecart_Nota.NOTA = Nota.NOTA "
Sql = Sql & "WHERE Nota.Id_IndiceProjet= " & MyIndice & "   "
Sql = Sql & "AND Temp_" & NmJob & "_Ecart_Nota.NUMNOTA Is Null;"

Debug.Print Sql

Set RsAjouter = Con.OpenRecordSet(Sql)


Sql = "SELECT '' AS [Avant/Après],Temp_" & NmJob & "_Ecart_Nota.ACTIVER, Temp_" & NmJob & "_Ecart_Nota.NOTA, Temp_" & NmJob & "_Ecart_Nota.NUMNOTA,Temp_" & NmJob & "_Ecart_Nota.[OPTION], '' AS Expr1,Nota.ACTIVER, Nota.NOTA, Nota.NUMNOTA,Nota.[OPTION]  "
Sql = Sql & "FROM Nota INNER JOIN Temp_" & NmJob & "_Ecart_Nota ON Nota.NOTA = Temp_" & NmJob & "_Ecart_Nota.NOTA "
Sql = Sql & "WHERE Nota.Id_IndiceProjet=" & MyIndice & ";"

Debug.Print Sql


Set RsModifier = Con.OpenRecordSet(Sql)

Txt = "REFF : " & TableProjet(1, I + 1)
Txt = Txt & Chr(10) & "DesCription : " & Chr(10) & Replace("" & TableProjet(2, I + 1), Chr(13), "")

Txt = Txt & Chr(10) & Chr(10) & "PIECE : " & Chr(10) & TableProjet(3, I + 1) & " -> " & TableProjet(3, I)
Txt = Txt & Chr(10) & Chr(10) & "PLAN : " & Chr(10) & TableProjet(4, I + 1) & " -> " & TableProjet(4, I)
Txt = Txt & Chr(10) & Chr(10) & "OUTIL : " & Chr(10) & TableProjet(5, I + 1) & " -> " & TableProjet(5, I)
Txt = Txt & Chr(10) & Chr(10) & "LISTE : " & Chr(10) & TableProjet(6, I + 1) & " -> " & TableProjet(6, I)
Txt = Txt & Chr(10) & Chr(10) & Chr(10) & Chr(10) & Chr(10) & Chr(10)
Debug.Print Txt
L = 0: C = 0
If MajEcartExcel(Val(TableProjet(0, LBound(TableProjet, 2))), MyWorkbook, L, C, RsSuprimer, RsAjouter, RsModifier, "Notas_Ecart", Txt) = True Then boolSave = True
Sql = "DROP TABLE Temp_" & NmJob & "_Composants_Ecart;"
Con.Execute Sql

Sql = "SELECT Composants.* INTO Temp_" & NmJob & "_Composants_Ecart "
Sql = Sql & "FROM Composants "
Sql = Sql & "WHERE Composants.Id_IndiceProjet= " & MyIndceMoins1 & ";"
Con.Execute Sql


Sql = "SELECT Temp_" & NmJob & "_Composants_Ecart.ACTIVER, Temp_" & NmJob & "_Composants_Ecart.DESIGNCOMP, Temp_" & NmJob & "_Composants_Ecart.NUMCOMP,  "
Sql = Sql & "Temp_" & NmJob & "_Composants_Ecart.REFCOMP, Temp_" & NmJob & "_Composants_Ecart.Path, Temp_" & NmJob & "_Composants_Ecart.OPTION "
Sql = Sql & "FROM Temp_" & NmJob & "_Composants_Ecart LEFT JOIN Composants ON Temp_" & NmJob & "_Composants_Ecart.NUMCOMP = Composants.NUMCOMP "
Sql = Sql & "WHERE Composants.NUMCOMP Is Null AND Composants.Id_IndiceProjet=" & MyIndice & ";"

Debug.Print Sql

Set RsSuprimer = Con.OpenRecordSet(Sql)

Sql = "SELECT Composants.ACTIVER,Composants.DESIGNCOMP, Composants.NUMCOMP, Composants.REFCOMP, Composants.Path,Composants.[OPTION] "
Sql = Sql & "FROM Composants LEFT JOIN Temp_" & NmJob & "_Composants_Ecart ON Composants.NUMCOMP = Temp_" & NmJob & "_Composants_Ecart.NUMCOMP "
Sql = Sql & "WHERE Composants.Id_IndiceProjet=" & MyIndice & " "
Sql = Sql & "AND Temp_" & NmJob & "_Composants_Ecart.NUMCOMP Is Null;"

Debug.Print Sql
Set RsAjouter = Con.OpenRecordSet(Sql)

Sql = "SELECT '' AS [Avant/Après], Temp_" & NmJob & "_Composants_Ecart.ACTIVER, Temp_" & NmJob & "_Composants_Ecart.DESIGNCOMP, "
Sql = Sql & "Temp_" & NmJob & "_Composants_Ecart.NUMCOMP, Temp_" & NmJob & "_Composants_Ecart.REFCOMP, Temp_" & NmJob & "_Composants_Ecart.Path, "
Sql = Sql & "Temp_" & NmJob & "_Composants_Ecart.OPTION, '' AS Expr1, Composants.ACTIVER, Composants.DESIGNCOMP, Composants.NUMCOMP, "
Sql = Sql & "Composants.REFCOMP, Composants.Path, Composants.OPTION "
Sql = Sql & "FROM Temp_" & NmJob & "_Composants_Ecart INNER JOIN Composants ON Temp_" & NmJob & "_Composants_Ecart.NUMCOMP = Composants.NUMCOMP "
'Sql = Sql & "ON Composants.NUMCOMP = Temp_" & NmJob & "_Composants_Ecart.NUMCOMP "
Sql = Sql & "WHERE Composants.Id_IndiceProjet=" & MyIndice & ";"


Debug.Print Sql

Set RsModifier = Con.OpenRecordSet(Sql)

L = 0: C = 0
If MajEcartExcel(Val(TableProjet(0, LBound(TableProjet, 2))), MyWorkbook, L, C, RsSuprimer, RsAjouter, RsModifier, "Composants_Ecart", Txt) = True Then boolSave = True


Sql = "DROP TABLE Temp_" & NmJob & "_T_Noeuds_Ecart;"
Con.Execute Sql
Sql = "SELECT T_Noeuds.* INTO Temp_" & NmJob & "_T_Noeuds_Ecart "
Sql = Sql & "FROM T_Noeuds "
Sql = Sql & "WHERE T_Noeuds.Id_IndiceProjet=" & MyIndceMoins1 & ";"
Debug.Print Sql
Con.Execute Sql

Sql = "SELECT Temp_" & NmJob & "_T_Noeuds_Ecart.ACTIVER, Temp_" & NmJob & "_T_Noeuds_Ecart.Fleche_Droite, Temp_" & NmJob & "_T_Noeuds_Ecart.TORON_PRINCIPAL,  "
Sql = Sql & "T_Noeuds.NŒUDS, Temp_" & NmJob & "_T_Noeuds_Ecart.LONGUEUR, Temp_" & NmJob & "_T_Noeuds_Ecart.LONGUEUR_CUMULEE,  "
Sql = Sql & "Temp_" & NmJob & "_T_Noeuds_Ecart.DESIGN_HAB, Temp_" & NmJob & "_T_Noeuds_Ecart.CODE_RSA, Temp_" & NmJob & "_T_Noeuds_Ecart.CODE_PSA,  "
Sql = Sql & "Temp_" & NmJob & "_T_Noeuds_Ecart.CODE_ENC, Temp_" & NmJob & "_T_Noeuds_Ecart.DIAMETRE, Temp_" & NmJob & "_T_Noeuds_Ecart.CLASSE_T,  "
Sql = Sql & "Temp_" & NmJob & "_T_Noeuds_Ecart.OPTION "
Sql = Sql & "FROM Temp_" & NmJob & "_T_Noeuds_Ecart LEFT JOIN T_Noeuds ON Temp_" & NmJob & "_T_Noeuds_Ecart.NŒUDS = T_Noeuds.NŒUDS "
Sql = Sql & "WHERE T_Noeuds.NŒUDS Is Null AND T_Noeuds.Id_IndiceProjet=" & MyIndice & ";"


Debug.Print Sql

Set RsSuprimer = Con.OpenRecordSet(Sql)

Sql = "SELECT T_Noeuds.ACTIVER,T_Noeuds.Fleche_Droite, T_Noeuds.TORON_PRINCIPAL,  T_Noeuds.NŒUDS,  "
Sql = Sql & "T_Noeuds.LONGUEUR, T_Noeuds.LONGUEUR_CUMULEE, T_Noeuds.DESIGN_HAB, T_Noeuds.CODE_RSA,  "
Sql = Sql & "T_Noeuds.CODE_PSA, T_Noeuds.CODE_ENC, T_Noeuds.DIAMETRE, T_Noeuds.CLASSE_T,T_Noeuds.[OPTION] "
Sql = Sql & "FROM T_Noeuds LEFT JOIN Temp_" & NmJob & "_T_Noeuds_Ecart ON T_Noeuds.NŒUDS = Temp_" & NmJob & "_T_Noeuds_Ecart.NŒUDS "
Sql = Sql & "WHERE T_Noeuds.Id_IndiceProjet=" & MyIndice & " "
Sql = Sql & "AND Temp_" & NmJob & "_T_Noeuds_Ecart.NŒUDS Is Null;"

Debug.Print Sql
Set RsAjouter = Con.OpenRecordSet(Sql)

Sql = "SELECT '' AS [Avant/Après],Temp_" & NmJob & "_T_Noeuds_Ecart.ACTIVER, Temp_" & NmJob & "_T_Noeuds_Ecart.Fleche_Droite, Temp_" & NmJob & "_T_Noeuds_Ecart.TORON_PRINCIPAL,  "
Sql = Sql & " Temp_" & NmJob & "_T_Noeuds_Ecart.NŒUDS, Temp_" & NmJob & "_T_Noeuds_Ecart.LONGUEUR,  "
Sql = Sql & "Temp_" & NmJob & "_T_Noeuds_Ecart.LONGUEUR_CUMULEE, Temp_" & NmJob & "_T_Noeuds_Ecart.DESIGN_HAB, Temp_" & NmJob & "_T_Noeuds_Ecart.CODE_RSA,  "
Sql = Sql & "Temp_" & NmJob & "_T_Noeuds_Ecart.CODE_PSA, Temp_" & NmJob & "_T_Noeuds_Ecart.CODE_ENC, Temp_" & NmJob & "_T_Noeuds_Ecart.DIAMETRE,  "
Sql = Sql & "Temp_" & NmJob & "_T_Noeuds_Ecart.CLASSE_T,Temp_" & NmJob & "_T_Noeuds_Ecart.[OPTION] , '' AS Expr1,T_Noeuds.ACTIVER, T_Noeuds.Fleche_Droite, T_Noeuds.TORON_PRINCIPAL,  "
Sql = Sql & " T_Noeuds.NŒUDS, T_Noeuds.LONGUEUR, T_Noeuds.LONGUEUR_CUMULEE, T_Noeuds.DESIGN_HAB,  "
Sql = Sql & "T_Noeuds.CODE_RSA, T_Noeuds.CODE_PSA, T_Noeuds.CODE_ENC, T_Noeuds.DIAMETRE, T_Noeuds.CLASSE_T,T_Noeuds.[OPTION] "
Sql = Sql & "FROM T_Noeuds INNER JOIN Temp_" & NmJob & "_T_Noeuds_Ecart ON T_Noeuds.NŒUDS = Temp_" & NmJob & "_T_Noeuds_Ecart.NŒUDS "
Sql = Sql & "WHERE T_Noeuds.Id_IndiceProjet=" & MyIndice & " ;"

Debug.Print Sql

Set RsModifier = Con.OpenRecordSet(Sql)
L = 0: C = 0
If MajEcartExcel(Val(TableProjet(0, LBound(TableProjet, 2))), MyWorkbook, L, C, RsSuprimer, RsAjouter, RsModifier, "Noeuds_Ecart", Txt) = True Then boolSave = True

Sql = "DROP TABLE Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart;"
Con.Execute Sql

Sql = "SELECT  Ligne_Tableau_fils.* INTO Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart "
Sql = Sql & "FROM Ligne_Tableau_fils "
Sql = Sql & "WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & MyIndceMoins1 & ";"
Con.Execute Sql

Sql = "SELECT Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.ACTIVER, Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.LIAI,  "
Sql = Sql & "Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.DESIGNATION, Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.FIL,  "
Sql = Sql & "Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.SECT, Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.TEINT,  "
Sql = Sql & "Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.TEINT2, Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.ISO,  "
Sql = Sql & "Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.LONG, Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.[LONG CP],  "
Sql = Sql & "Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.COUPE, Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.POS, Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.[POS-OUT],  "
Sql = Sql & "Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.FA, Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.APP, Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.VOI,  "
Sql = Sql & "Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.POS2, Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.[POS-OUT2], Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.FA2,  "
Sql = Sql & "Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.APP2, Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.VOI2, Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.PRECOG,  "
Sql = Sql & "Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.OPTION "
Sql = Sql & "FROM Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart LEFT JOIN Ligne_Tableau_fils ON (Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.VOI2 = Ligne_Tableau_fils.VOI2) AND (Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.VOI = Ligne_Tableau_fils.VOI) AND (Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.LIAI = Ligne_Tableau_fils.LIAI) AND (Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.APP = Ligne_Tableau_fils.APP) AND (Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.APP2 = Ligne_Tableau_fils.APP2)  "
'Sql = Sql & "ON Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.FIL = Ligne_Tableau_fils.FIL "
Sql = Sql & "WHERE Ligne_Tableau_fils.FIL Is Null AND Ligne_Tableau_fils.Id_IndiceProjet=" & MyIndice & ";"

Debug.Print Sql


Set RsSuprimer = Con.OpenRecordSet(Sql)

Sql = "SELECT  Ligne_Tableau_fils.ACTIVER,Ligne_Tableau_fils.LIAI, Ligne_Tableau_fils.DESIGNATION, Ligne_Tableau_fils.FIL,  "
Sql = Sql & "Ligne_Tableau_fils.SECT, Ligne_Tableau_fils.TEINT, Ligne_Tableau_fils.TEINT2,  "
Sql = Sql & "Ligne_Tableau_fils.ISO, Ligne_Tableau_fils.LONG, Ligne_Tableau_fils.[LONG CP],  "
Sql = Sql & "Ligne_Tableau_fils.COUPE, Ligne_Tableau_fils.POS, Ligne_Tableau_fils.[POS-OUT],  "
Sql = Sql & "Ligne_Tableau_fils.FA, Ligne_Tableau_fils.APP, Ligne_Tableau_fils.VOI, Ligne_Tableau_fils.POS2,  "
Sql = Sql & "Ligne_Tableau_fils.[POS-OUT2], Ligne_Tableau_fils.FA2, Ligne_Tableau_fils.APP2,  "
Sql = Sql & "Ligne_Tableau_fils.VOI2, Ligne_Tableau_fils.PRECOG, Ligne_Tableau_fils.OPTION "
Sql = Sql & "FROM Ligne_Tableau_fils LEFT JOIN Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart ON (Ligne_Tableau_fils.VOI2 = Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.VOI2) AND (Ligne_Tableau_fils.VOI = Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.VOI) AND (Ligne_Tableau_fils.LIAI = Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.LIAI) AND (Ligne_Tableau_fils.APP = Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.APP) AND (Ligne_Tableau_fils.APP2 = Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.APP2) "
Sql = Sql & "WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & MyIndice & "  "
Sql = Sql & "AND Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.LIAI Is Null;"

Debug.Print Sql
Set RsAjouter = Con.OpenRecordSet(Sql)

Sql = "SELECT '' AS [Avant/Après],Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.ACTIVER, Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.LIAI, Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.DESIGNATION,   "
Sql = Sql & "Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.FIL, Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.SECT, Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.TEINT,   "
Sql = Sql & "Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.TEINT2, Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.ISO, Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.LONG,   "
Sql = Sql & "Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.[LONG CP], Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.COUPE, Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.POS,   "
Sql = Sql & "Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.[POS-OUT], Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.FA, Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.APP,   "
Sql = Sql & "Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.VOI, Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.POS2, Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.[POS-OUT2],   "
Sql = Sql & "Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.FA2, Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.APP2, Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.VOI2,   "
Sql = Sql & "Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.PRECOG, Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.OPTION, '' AS Expr1,Ligne_Tableau_fils.ACTIVER, Ligne_Tableau_fils.LIAI,   "
Sql = Sql & "Ligne_Tableau_fils.DESIGNATION, Ligne_Tableau_fils.FIL, Ligne_Tableau_fils.SECT, Ligne_Tableau_fils.TEINT,   "
Sql = Sql & "Ligne_Tableau_fils.TEINT2, Ligne_Tableau_fils.ISO , Ligne_Tableau_fils.Long, Ligne_Tableau_fils.[LONG CP],  "
Sql = Sql & " Ligne_Tableau_fils.Coupe, Ligne_Tableau_fils.POS, Ligne_Tableau_fils.[POS-OUT], Ligne_Tableau_fils.FA,   "
Sql = Sql & "Ligne_Tableau_fils.App, Ligne_Tableau_fils.VOI, Ligne_Tableau_fils.POS2, Ligne_Tableau_fils.[POS-OUT2],   "
Sql = Sql & "Ligne_Tableau_fils.FA2, Ligne_Tableau_fils.APP2, Ligne_Tableau_fils.VOI2, Ligne_Tableau_fils.PRECOG, Ligne_Tableau_fils.Option  "
Sql = Sql & "FROM Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart INNER JOIN Ligne_Tableau_fils ON (Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.VOI2 = Ligne_Tableau_fils.VOI2) AND (Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.VOI = Ligne_Tableau_fils.VOI) AND (Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.LIAI = Ligne_Tableau_fils.LIAI) AND (Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.APP = Ligne_Tableau_fils.APP) AND (Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.APP2 = Ligne_Tableau_fils.APP2)  "
Sql = Sql & "WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & MyIndice & " ;"

Debug.Print Sql

Set RsModifier = Con.OpenRecordSet(Sql)

L = 0: C = 0
If MajEcartExcel(Val(TableProjet(0, LBound(TableProjet, 2))), MyWorkbook, L, C, RsSuprimer, RsAjouter, RsModifier, "Tableau_fils_Ecart", Txt) = True Then boolSave = True

Sql = "DROP TABLE Temp_" & NmJob & "_Connecteurs_Ecart;"
Con.Execute Sql

Sql = " SELECT  Connecteurs.* INTO Temp_" & NmJob & "_Connecteurs_Ecart "
Sql = Sql & "FROM Connecteurs "
Sql = Sql & "WHERE Connecteurs.Id_IndiceProjet=" & MyIndceMoins1 & " ;"
Con.Execute Sql


Sql = "SELECT Temp_" & NmJob & "_Connecteurs_Ecart.ACTIVER, Temp_" & NmJob & "_Connecteurs_Ecart.CONNECTEUR,  "
Sql = Sql & "Temp_" & NmJob & "_Connecteurs_Ecart.RefConnecteurFour, Temp_" & NmJob & "_Connecteurs_Ecart.[O/N],  "
Sql = Sql & "Temp_" & NmJob & "_Connecteurs_Ecart.DESIGNATION, Temp_" & NmJob & "_Connecteurs_Ecart.CODE_APP,  "
Sql = Sql & "Temp_" & NmJob & "_Connecteurs_Ecart.N°, Temp_" & NmJob & "_Connecteurs_Ecart.POS, Temp_" & NmJob & "_Connecteurs_Ecart.[POS-OUT],  "
Sql = Sql & "Temp_" & NmJob & "_Connecteurs_Ecart.PRECO1, Temp_" & NmJob & "_Connecteurs_Ecart.PRECO2, Temp_" & NmJob & "_Connecteurs_Ecart.[100%],  "
Sql = Sql & "Temp_" & NmJob & "_Connecteurs_Ecart.OPTION "
Sql = Sql & "FROM Temp_" & NmJob & "_Connecteurs_Ecart LEFT JOIN Connecteurs ON Temp_" & NmJob & "_Connecteurs_Ecart.CODE_APP = Connecteurs.CODE_APP "
Sql = Sql & "WHERE Connecteurs.Id_IndiceProjet=" & MyIndice & " AND Connecteurs.CODE_APP Is Null;"


Debug.Print Sql

Set RsSuprimer = Con.OpenRecordSet(Sql)

Sql = "SELECT Connecteurs.ACTIVER,Connecteurs.CONNECTEUR, Connecteurs.RefConnecteurFour,Connecteurs.[O/N], Connecteurs.DESIGNATION, Connecteurs.CODE_APP,   "
Sql = Sql & "Connecteurs.N°, Connecteurs.POS, Connecteurs.[POS-OUT], Connecteurs.PRECO1, Connecteurs.PRECO2,   "
Sql = Sql & "Connecteurs.[100%], Connecteurs.OPTION  "
Sql = Sql & "FROM Connecteurs LEFT JOIN Temp_" & NmJob & "_Connecteurs_Ecart ON Connecteurs.CODE_APP = Temp_" & NmJob & "_Connecteurs_Ecart.CODE_APP "
Sql = Sql & "WHERE Connecteurs.Id_IndiceProjet=" & MyIndice & " "
Sql = Sql & "AND Temp_" & NmJob & "_Connecteurs_Ecart.N° Is Null;"

Debug.Print Sql

Set RsAjouter = Con.OpenRecordSet(Sql)

Sql = "SELECT '' AS [Avant/Après], Temp_" & NmJob & "_Connecteurs_Ecart.ACTIVER, Temp_" & NmJob & "_Connecteurs_Ecart.CONNECTEUR,  "
Sql = Sql & "Temp_" & NmJob & "_Connecteurs_Ecart.RefConnecteurFour, Temp_" & NmJob & "_Connecteurs_Ecart.[O/N], Temp_" & NmJob & "_Connecteurs_Ecart.DESIGNATION,  "
Sql = Sql & "Temp_" & NmJob & "_Connecteurs_Ecart.CODE_APP, Temp_" & NmJob & "_Connecteurs_Ecart.N°, Temp_" & NmJob & "_Connecteurs_Ecart.POS,  "
Sql = Sql & "Temp_" & NmJob & "_Connecteurs_Ecart.[POS-OUT], Temp_" & NmJob & "_Connecteurs_Ecart.PRECO1, Temp_" & NmJob & "_Connecteurs_Ecart.PRECO2,  "
Sql = Sql & "Temp_" & NmJob & "_Connecteurs_Ecart.[100%], Temp_" & NmJob & "_Connecteurs_Ecart.OPTION, '' AS Expr1, Connecteurs.ACTIVER,  "
Sql = Sql & "Connecteurs.CONNECTEUR, Connecteurs.RefConnecteurFour, Connecteurs.[O/N], Connecteurs.DESIGNATION, Connecteurs.CODE_APP,  "
Sql = Sql & "Connecteurs.N°, Connecteurs.POS, Connecteurs.[POS-OUT], Connecteurs.PRECO1, Connecteurs.PRECO2, Connecteurs.[100%],  "
Sql = Sql & "Connecteurs.OPTION "
Sql = Sql & "FROM Connecteurs INNER JOIN Temp_" & NmJob & "_Connecteurs_Ecart ON Connecteurs.CODE_APP = Temp_" & NmJob & "_Connecteurs_Ecart.CODE_APP "
Sql = Sql & "WHERE Connecteurs.Id_IndiceProjet=" & MyIndice & ";"

 

Debug.Print Sql

Set RsModifier = Con.OpenRecordSet(Sql)

L = 0: C = 0
If MajEcartExcel(Val(TableProjet(0, LBound(TableProjet, 2))), MyWorkbook, L, C, RsSuprimer, RsAjouter, RsModifier, "Connecteurs_Ecart", Txt) = True Then boolSave = True

Sql = "DROP TABLE Temp_" & NmJob & "_T_Critères_Ecart;"
Con.Execute Sql

Sql = "SELECT  T_Critères.* INTO Temp_" & NmJob & "_T_Critères_Ecart "
Sql = Sql & "FROM T_Critères "
Sql = Sql & "WHERE T_Critères.Id_IndiceProjet=" & MyIndceMoins1 & ";"

Con.Execute Sql


Sql = "SELECT Temp_" & NmJob & "_T_Critères_Ecart.ACTIVER, Temp_" & NmJob & "_T_Critères_Ecart.CODE_CRITERE, Temp_" & NmJob & "_T_Critères_Ecart.CRITERES "
Sql = Sql & "FROM Temp_" & NmJob & "_T_Critères_Ecart LEFT JOIN T_Critères ON Temp_" & NmJob & "_T_Critères_Ecart.CODE_CRITERE = T_Critères.CODE_CRITERE "
Sql = Sql & "WHERE T_Critères.Id_IndiceProjet=" & MyIndice & "  AND T_Critères.CODE_CRITERE Is Null;"


Debug.Print Sql

Set RsSuprimer = Con.OpenRecordSet(Sql)

Sql = "SELECT T_Critères.ACTIVER,T_Critères.CODE_CRITERE, T_Critères.CRITERES  "
Sql = Sql & "FROM T_Critères LEFT JOIN Temp_" & NmJob & "_T_Critères_Ecart ON (T_Critères.CODE_CRITERE = Temp_" & NmJob & "_T_Critères_Ecart.CODE_CRITERE)    "
Sql = Sql & "WHERE T_Critères.Id_IndiceProjet=" & MyIndice & "  AND Temp_" & NmJob & "_T_Critères_Ecart.CODE_CRITERE Is Null;"



Debug.Print Sql

Set RsAjouter = Con.OpenRecordSet(Sql)

Sql = "SELECT '' AS [Avant/Après],Temp_" & NmJob & "_T_Critères_Ecart.ACTIVER, Temp_" & NmJob & "_T_Critères_Ecart.CODE_CRITERE,  "
Sql = Sql & "Temp_" & NmJob & "_T_Critères_Ecart.CRITERES, '' AS Expr1, T_Critères.ACTIVER, T_Critères.CODE_CRITERE,  "
Sql = Sql & "T_Critères.CRITERES  "
Sql = Sql & "FROM Temp_" & NmJob & "_T_Critères_Ecart INNER JOIN T_Critères ON Temp_" & NmJob & "_T_Critères_Ecart.CODE_CRITERE = T_Critères.CODE_CRITERE    "
Sql = Sql & "Where T_Critères.Id_IndiceProjet = " & MyIndice & "  ORDER BY T_Critères.CODE_CRITERE;"



Debug.Print Sql

Set RsModifier = Con.OpenRecordSet(Sql)

L = 0: C = 0
If MajEcartExcel(Val(TableProjet(0, LBound(TableProjet, 2))), MyWorkbook, L, C, RsSuprimer, RsAjouter, RsModifier, "Critères_Ecart", Txt) = True Then boolSave = True
Next
Set RsAjouter = Con.CloseRecordSet(RsAjouter)
Set RsModifier = Con.CloseRecordSet(RsModifier)
Set RsSuprimer = Con.CloseRecordSet(RsSuprimer)


Sql = "DROP TABLE Temp_" & NmJob & "_Ecart_Nota;"
Con.Execute Sql

Sql = "DROP TABLE Temp_" & NmJob & "_Composants_Ecart;"
Con.Execute Sql

Sql = "DROP TABLE Temp_" & NmJob & "_T_Noeuds_Ecart;"
Con.Execute Sql

Sql = "DROP TABLE Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart;"
Con.Execute Sql

Sql = "DROP TABLE Temp_" & NmJob & "_Connecteurs_Ecart;"
Con.Execute Sql

Sql = "DROP TABLE Temp_" & NmJob & "_T_Critères_Ecart;"
Con.Execute Sql

If boolSave = True Then
Sql = "SELECT RqCartouche.* "
Sql = Sql & "FROM RqCartouche "
Sql = Sql & "WHERE T_indiceProjet.Id=" & MyIndice & " and T_indiceProjet.DNC IS NOT NULL;"
Set RsModifier = Con.OpenRecordSet(Sql)
    If RsModifier.EOF = False Then
    
        PathPl = PathArchive(TableauPath("PathArchiveAutocad"), "" & RsModifier!Client, "" & RsModifier!CleAc, "" & RsModifier!Pieces, "LIEC", RsModifier.Fields("LIEC"), MyIndice, RsModifier.Fields("PI_Indice"), "", RsModifier!Version, True)
       If Fso.FileExists(PathPl & ".xls") = True Then Fso.DeleteFile PathPl & ".xls"
        MyWorkbook.SaveAs PathPl, ReadOnlyRecommended:=True
    
        If IdFils <> 0 Then
            Sql = "SELECT RqCartouche.* "
            Sql = Sql & "FROM RqCartouche "
            Sql = Sql & "WHERE T_indiceProjet.Id=" & IdFils & ";"
            Set RsModifier = Con.OpenRecordSet(Sql)
            PathPl2 = PathArchive(PathArchiveAutocad, "" & RsModifier!Client, "" & RsModifier!CleAc, "" & RsModifier!Pieces, "LIEC", RsModifier.Fields("LIEC"), RsModifier!Id, RsModifier.Fields("PI_Indice"), "", RsModifier!Version, True)
            
            Racourci "" & PathPl2, "" & PathPl, "XLS"
        End If
       SubActionCorrective MyIndice, IdFils
    End If
End If


Sortie:

MyWorkbook.Close False
MyExcel.Quit
End Sub
