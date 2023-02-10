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
Dim MyEcel As New EXCEL.Application
boolSave = False
   PathArchiveAutocad = TableauPath.Item("PathArchiveAutocad")
     If Left(PathArchiveAutocad, 2) <> "\\" And Left(PathArchiveAutocad, 1) = "\" Then PathArchiveAutocad = TableauPath.Item("PathServer") & PathArchiveAutocad
  If Right(PathArchiveAutocad, 2) = "\\" Then PathArchiveAutocad = Mid(PathArchiveAutocad, 1, Len(MyPath) - 1)
'MyEcel.Visible = True
Set MyWorkbook = MyEcel.Workbooks.Add

Sql = "SELECT T_indiceProjet_1.Id, T_indiceProjet_1.ReffIndice, T_indiceProjet_1.Description,  "
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
Sql = Sql & "[T_indiceProjet_1].[Li] & '_' & Trim('' & [T_indiceProjet_1].[Li_Indice]) "
Sql = Sql & "ORDER BY T_indiceProjet_1.Id;"
Set Rs = Con.OpenRecordSet(Sql)
TableProjet = Rs.GetRows
If UBound(TableProjet, 2) - 1 = -1 Then GoTo Sortie
For i = 0 To UBound(TableProjet, 2) - 1

    Sql = "DROP TABLE Temp_" & NmJob & "_Ecart_Nota;"
    Con.Exequte Sql
    Sql = "SELECT Nota.Id, Nota.Id_IndiceProjet, " & TableProjet(0, i + 1) & " AS Id_IndiceProjet2, "
    Sql = Sql & "Nota.ACTIVER,Nota.NOTA, Nota.NUMNOTA INTO Temp_" & NmJob & "_Ecart_Nota "
    Sql = Sql & "FROM Nota "
    Sql = Sql & "WHERE Nota.Id_IndiceProjet=" & TableProjet(0, i) & ";"
    
    Con.Exequte Sql
    
    Sql = "SELECT Temp_" & NmJob & "_Ecart_Nota.ACTIVER,Temp_" & NmJob & "_Ecart_Nota.NOTA, Temp_" & NmJob & "_Ecart_Nota.NUMNOTA "
    Sql = Sql & "FROM Temp_" & NmJob & "_Ecart_Nota LEFT JOIN Nota ON (Temp_" & NmJob & "_Ecart_Nota.NUMNOTA = Nota.NUMNOTA)  "
    Sql = Sql & "AND (Temp_" & NmJob & "_Ecart_Nota.Id_IndiceProjet2 = Nota.Id_IndiceProjet) "
    Sql = Sql & "WHERE Temp_" & NmJob & "_Ecart_Nota.Id_IndiceProjet2= " & TableProjet(0, i + 1) & "   "
    Sql = Sql & "AND Nota.NUMNOTA Is Null;"

    Debug.Print Sql


Set RsSuprimer = Con.OpenRecordSet(Sql)

Sql = "SELECT Nota.ACTIVER,Nota.NOTA, Nota.NUMNOTA "
Sql = Sql & "FROM Nota LEFT JOIN Temp_" & NmJob & "_Ecart_Nota ON (Nota.NUMNOTA = Temp_" & NmJob & "_Ecart_Nota.NUMNOTA)  "
Sql = Sql & "AND (Nota.Id_IndiceProjet = Temp_" & NmJob & "_Ecart_Nota.Id_IndiceProjet) "
Sql = Sql & "WHERE Nota.Id_IndiceProjet= " & TableProjet(0, i + 1) & "   "
Sql = Sql & "AND Temp_" & NmJob & "_Ecart_Nota.NUMNOTA Is Null;"

Debug.Print Sql

Set RsAjouter = Con.OpenRecordSet(Sql)


Sql = "SELECT '' AS [Avant/Après],Temp_" & NmJob & "_Ecart_Nota.ACTIVER, Temp_" & NmJob & "_Ecart_Nota.NOTA, Temp_" & NmJob & "_Ecart_Nota.NUMNOTA, '' AS Expr1,Nota.ACTIVER, Nota.NOTA, Nota.NUMNOTA  "
Sql = Sql & "FROM Temp_" & NmJob & "_Ecart_Nota INNER JOIN Nota ON (Temp_" & NmJob & "_Ecart_Nota.NUMNOTA = Nota.NUMNOTA)   "
Sql = Sql & "AND (Temp_" & NmJob & "_Ecart_Nota.Id_IndiceProjet2 = Nota.Id_IndiceProjet)  "
Sql = Sql & "WHERE Nota.Id_IndiceProjet=" & TableProjet(0, i + 1) & ";"

Debug.Print Sql


Set RsModifier = Con.OpenRecordSet(Sql)

txt = "REFF : " & TableProjet(1, i + 1)
txt = txt & Chr(10) & "DesCription : " & Chr(10) & Replace(TableProjet(2, i + 1), Chr(13), "")

txt = txt & Chr(10) & Chr(10) & "PIECE : " & Chr(10) & TableProjet(3, i) & " -> " & TableProjet(3, i + 1)
txt = txt & Chr(10) & Chr(10) & "PLAN : " & Chr(10) & TableProjet(4, i) & " -> " & TableProjet(4, i + 1)
txt = txt & Chr(10) & Chr(10) & "OUTIL : " & Chr(10) & TableProjet(5, i) & " -> " & TableProjet(5, i + 1)
txt = txt & Chr(10) & Chr(10) & "LISTE : " & Chr(10) & TableProjet(6, i) & " -> " & TableProjet(6, i + 1)
txt = txt & Chr(10) & Chr(10) & Chr(10) & Chr(10) & Chr(10) & Chr(10)
Debug.Print txt
L = 0: C = 0
If MajEcartExcel(Val(TableProjet(0, UBound(TableProjet, 2))), MyWorkbook, L, C, RsSuprimer, RsAjouter, RsModifier, "Notas_Ecart", txt) = True Then boolSave = True
Sql = "DROP TABLE Temp_" & NmJob & "_Composants_Ecart;"
Con.Exequte Sql

Sql = "SELECT Composants.ACTIVER,Composants.Id, Composants.Id_IndiceProjet, " & TableProjet(0, i + 1) & " AS Id_IndiceProjet2, Composants.DESIGNCOMP,  "
Sql = Sql & "Composants.NUMCOMP, Composants.REFCOMP, Composants.Path INTO Temp_" & NmJob & "_Composants_Ecart "
Sql = Sql & "FROM Composants "
Sql = Sql & "WHERE Composants.Id_IndiceProjet= " & TableProjet(0, i) & ";"
Con.Exequte Sql


Sql = "SELECT Temp_" & NmJob & "_Composants_Ecart.ACTIVER,Temp_" & NmJob & "_Composants_Ecart.DESIGNCOMP, Temp_" & NmJob & "_Composants_Ecart.NUMCOMP,  "
Sql = Sql & "Temp_" & NmJob & "_Composants_Ecart.REFCOMP, Temp_" & NmJob & "_Composants_Ecart.Path "
Sql = Sql & "FROM Temp_" & NmJob & "_Composants_Ecart LEFT JOIN Composants ON (Temp_" & NmJob & "_Composants_Ecart.NUMCOMP = Composants.NUMCOMP)  "
Sql = Sql & "AND (Temp_" & NmJob & "_Composants_Ecart.Id_IndiceProjet2 = Composants.Id_IndiceProjet) "
Sql = Sql & "WHERE Composants.NUMCOMP Is Null  "
Sql = Sql & "AND Temp_" & NmJob & "_Composants_Ecart.Id_IndiceProjet=" & TableProjet(0, i + 1) & ";"

Debug.Print Sql

Set RsSuprimer = Con.OpenRecordSet(Sql)

Sql = "SELECT Composants.ACTIVER,Composants.DESIGNCOMP, Composants.NUMCOMP, Composants.REFCOMP, Composants.Path "
Sql = Sql & "FROM Composants LEFT JOIN Temp_" & NmJob & "_Composants_Ecart ON (Composants.NUMCOMP ="
Sql = Sql & "Temp_" & NmJob & "_Composants_Ecart.NUMCOMP) AND (Composants.Id_IndiceProjet = Temp_" & NmJob & "_Composants_Ecart.Id_IndiceProjet2) "
Sql = Sql & "WHERE Composants.Id_IndiceProjet=" & TableProjet(0, i + 1) & " "
Sql = Sql & "AND Temp_" & NmJob & "_Composants_Ecart.NUMCOMP Is Null;"

Debug.Print Sql
Set RsAjouter = Con.OpenRecordSet(Sql)

Sql = "SELECT '' AS [Avant/Après],Temp_" & NmJob & "_Composants_Ecart.ACTIVER, Temp_" & NmJob & "_Composants_Ecart.DESIGNCOMP, Temp_" & NmJob & "_Composants_Ecart.NUMCOMP,  "
Sql = Sql & "Temp_" & NmJob & "_Composants_Ecart.REFCOMP, Temp_" & NmJob & "_Composants_Ecart.Path, '' AS Expr1,Composants.ACTIVER, Composants.DESIGNCOMP,  "
Sql = Sql & "Composants.NUMCOMP, Composants.REFCOMP, Composants.Path "
Sql = Sql & "FROM Composants INNER JOIN Temp_" & NmJob & "_Composants_Ecart ON (Composants.NUMCOMP = Temp_" & NmJob & "_Composants_Ecart.NUMCOMP)  "
Sql = Sql & "AND (Composants.Id_IndiceProjet = Temp_" & NmJob & "_Composants_Ecart.Id_IndiceProjet2) "
Sql = Sql & "WHERE Composants.Id_IndiceProjet=" & TableProjet(0, i + 1) & ";"

Debug.Print Sql

Set RsModifier = Con.OpenRecordSet(Sql)

L = 0: C = 0
If MajEcartExcel(Val(TableProjet(0, UBound(TableProjet, 2))), MyWorkbook, L, C, RsSuprimer, RsAjouter, RsModifier, "Composants_Ecart", txt) = True Then boolSave = True


Sql = "DROP TABLE Temp_" & NmJob & "_T_Noeuds_Ecart;"
Con.Exequte Sql
Sql = "SELECT T_Noeuds.*, " & TableProjet(0, i + 1) & " AS Id_IndiceProjet2 INTO Temp_" & NmJob & "_T_Noeuds_Ecart "
Sql = Sql & "FROM T_Noeuds "
Sql = Sql & "WHERE T_Noeuds.Id_IndiceProjet=" & TableProjet(0, i) & ";"

Con.Exequte Sql

Sql = "SELECT Temp_" & NmJob & "_T_Noeuds_Ecart.ACTIVER,Temp_" & NmJob & "_T_Noeuds_Ecart.Fleche_Droite, Temp_" & NmJob & "_T_Noeuds_Ecart.TORON_PRINCIPAL,  "
Sql = Sql & " T_Noeuds.NŒUDS, Temp_" & NmJob & "_T_Noeuds_Ecart.LONGUEUR,  "
Sql = Sql & "Temp_" & NmJob & "_T_Noeuds_Ecart.LONGUEUR_CUMULEE, Temp_" & NmJob & "_T_Noeuds_Ecart.DESIGN_HAB,  "
Sql = Sql & "Temp_" & NmJob & "_T_Noeuds_Ecart.CODE_RSA, Temp_" & NmJob & "_T_Noeuds_Ecart.CODE_PSA, Temp_" & NmJob & "_T_Noeuds_Ecart.CODE_ENC,  "
Sql = Sql & "Temp_" & NmJob & "_T_Noeuds_Ecart.DIAMETRE, Temp_" & NmJob & "_T_Noeuds_Ecart.CLASSE_T "
Sql = Sql & "FROM Temp_" & NmJob & "_T_Noeuds_Ecart LEFT JOIN T_Noeuds ON (Temp_" & NmJob & "_T_Noeuds_Ecart.Id_IndiceProjet2 =  "
Sql = Sql & "T_Noeuds.Id_IndiceProjet) AND (Temp_" & NmJob & "_T_Noeuds_Ecart.NŒUDS = T_Noeuds.NŒUDS) "
Sql = Sql & "WHERE T_Noeuds.NŒUDS Is Null "
Sql = Sql & "AND Temp_" & NmJob & "_T_Noeuds_Ecart.Id_IndiceProjet=" & TableProjet(0, i + 1) & ";"

Debug.Print Sql

Set RsSuprimer = Con.OpenRecordSet(Sql)

Sql = "SELECT T_Noeuds.ACTIVER,T_Noeuds.Fleche_Droite, T_Noeuds.TORON_PRINCIPAL,  T_Noeuds.NŒUDS,  "
Sql = Sql & "T_Noeuds.LONGUEUR, T_Noeuds.LONGUEUR_CUMULEE, T_Noeuds.DESIGN_HAB, T_Noeuds.CODE_RSA,  "
Sql = Sql & "T_Noeuds.CODE_PSA, T_Noeuds.CODE_ENC, T_Noeuds.DIAMETRE, T_Noeuds.CLASSE_T "
Sql = Sql & "FROM T_Noeuds LEFT JOIN Temp_" & NmJob & "_T_Noeuds_Ecart ON (T_Noeuds.Id_IndiceProjet =  "
Sql = Sql & "Temp_" & NmJob & "_T_Noeuds_Ecart.Id_IndiceProjet2) AND (T_Noeuds.NŒUDS = Temp_" & NmJob & "_T_Noeuds_Ecart.NŒUDS) "
Sql = Sql & "WHERE T_Noeuds.Id_IndiceProjet=" & TableProjet(0, i + 1) & " "
Sql = Sql & "AND Temp_" & NmJob & "_T_Noeuds_Ecart.NŒUDS Is Null;"

Debug.Print Sql
Set RsAjouter = Con.OpenRecordSet(Sql)

Sql = "SELECT '' AS [Avant/Après],Temp_" & NmJob & "_T_Noeuds_Ecart.ACTIVER, Temp_" & NmJob & "_T_Noeuds_Ecart.Fleche_Droite, Temp_" & NmJob & "_T_Noeuds_Ecart.TORON_PRINCIPAL,  "
Sql = Sql & " Temp_" & NmJob & "_T_Noeuds_Ecart.NŒUDS, Temp_" & NmJob & "_T_Noeuds_Ecart.LONGUEUR,  "
Sql = Sql & "Temp_" & NmJob & "_T_Noeuds_Ecart.LONGUEUR_CUMULEE, Temp_" & NmJob & "_T_Noeuds_Ecart.DESIGN_HAB, Temp_" & NmJob & "_T_Noeuds_Ecart.CODE_RSA,  "
Sql = Sql & "Temp_" & NmJob & "_T_Noeuds_Ecart.CODE_PSA, Temp_" & NmJob & "_T_Noeuds_Ecart.CODE_ENC, Temp_" & NmJob & "_T_Noeuds_Ecart.DIAMETRE,  "
Sql = Sql & "Temp_" & NmJob & "_T_Noeuds_Ecart.CLASSE_T, '' AS Expr1,T_Noeuds.ACTIVER, T_Noeuds.Fleche_Droite, T_Noeuds.TORON_PRINCIPAL,  "
Sql = Sql & " T_Noeuds.NŒUDS, T_Noeuds.LONGUEUR, T_Noeuds.LONGUEUR_CUMULEE, T_Noeuds.DESIGN_HAB,  "
Sql = Sql & "T_Noeuds.CODE_RSA, T_Noeuds.CODE_PSA, T_Noeuds.CODE_ENC, T_Noeuds.DIAMETRE, T_Noeuds.CLASSE_T "
Sql = Sql & "FROM T_Noeuds INNER JOIN Temp_" & NmJob & "_T_Noeuds_Ecart  "
Sql = Sql & "ON (T_Noeuds.Id_IndiceProjet = Temp_" & NmJob & "_T_Noeuds_Ecart.Id_IndiceProjet2)  "
Sql = Sql & "AND (T_Noeuds.NŒUDS = Temp_" & NmJob & "_T_Noeuds_Ecart.NŒUDS) "
Sql = Sql & "WHERE T_Noeuds.Id_IndiceProjet=" & TableProjet(0, i + 1) & ";"

Debug.Print Sql

Set RsModifier = Con.OpenRecordSet(Sql)
L = 0: C = 0
If MajEcartExcel(Val(TableProjet(0, UBound(TableProjet, 2))), MyWorkbook, L, C, RsSuprimer, RsAjouter, RsModifier, "Noeuds_Ecart", txt) = True Then boolSave = True

Sql = "DROP TABLE Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart;"
Con.Exequte Sql

Sql = "SELECT " & TableProjet(0, i + 1) & " AS Id_IndiceProjet2, Ligne_Tableau_fils.* "
Sql = Sql & "INTO Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart "
Sql = Sql & "FROM Ligne_Tableau_fils "
Sql = Sql & "WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & TableProjet(0, i) & ";"
Con.Exequte Sql

Sql = "SELECT Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.ACTIVER,Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.LIAI, Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.DESIGNATION,  "
Sql = Sql & "Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.FIL, Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.SECT,  "
Sql = Sql & "Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.TEINT, Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.TEINT2,  "
Sql = Sql & "Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.ISO, Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.LONG,  "
Sql = Sql & "Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.[LONG CP], Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.COUPE,  "
Sql = Sql & "Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.POS, Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.[POS-OUT],  "
Sql = Sql & "Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.FA, Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.APP,  "
Sql = Sql & "Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.VOI, Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.POS2,  "
Sql = Sql & "Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.[POS-OUT2], Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.FA2,  "
Sql = Sql & "Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.APP2, Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.VOI2,  "
Sql = Sql & "Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.PRECO, Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.OPTION "
Sql = Sql & "FROM Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart LEFT JOIN Ligne_Tableau_fils  "
Sql = Sql & "ON (Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.FIL = Ligne_Tableau_fils.FIL)  "
Sql = Sql & "AND (Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.Id_IndiceProjet2 = Ligne_Tableau_fils.Id_IndiceProjet) "
Sql = Sql & "WHERE Ligne_Tableau_fils.FIL Is Null  "
Sql = Sql & "AND Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.Id_IndiceProjet=" & TableProjet(0, i + 1) & ";"

Debug.Print Sql


Set RsSuprimer = Con.OpenRecordSet(Sql)

Sql = "SELECT Ligne_Tableau_Flis.ACTIVER,Ligne_Tableau_fils.LIAI, Ligne_Tableau_fils.DESIGNATION, Ligne_Tableau_fils.FIL,  "
Sql = Sql & "Ligne_Tableau_fils.SECT, Ligne_Tableau_fils.TEINT, Ligne_Tableau_fils.TEINT2,  "
Sql = Sql & "Ligne_Tableau_fils.ISO, Ligne_Tableau_fils.LONG, Ligne_Tableau_fils.[LONG CP],  "
Sql = Sql & "Ligne_Tableau_fils.COUPE, Ligne_Tableau_fils.POS, Ligne_Tableau_fils.[POS-OUT],  "
Sql = Sql & "Ligne_Tableau_fils.FA, Ligne_Tableau_fils.APP, Ligne_Tableau_fils.VOI, Ligne_Tableau_fils.POS2,  "
Sql = Sql & "Ligne_Tableau_fils.[POS-OUT2], Ligne_Tableau_fils.FA2, Ligne_Tableau_fils.APP2,  "
Sql = Sql & "Ligne_Tableau_fils.VOI2, Ligne_Tableau_fils.PRECO, Ligne_Tableau_fils.OPTION "
Sql = Sql & "FROM Ligne_Tableau_fils LEFT JOIN Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart  "
Sql = Sql & "ON (Ligne_Tableau_fils.Id_IndiceProjet = Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.Id_IndiceProjet2)  "
Sql = Sql & "AND (Ligne_Tableau_fils.FIL = Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.FIL) "
Sql = Sql & "WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & TableProjet(0, i + 1) & "  "
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
Sql = Sql & "Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.PRECO, Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.OPTION, '' AS Expr1,Ligne_Tableau_Flis.ACTIVER, Ligne_Tableau_fils.LIAI,   "
Sql = Sql & "Ligne_Tableau_fils.DESIGNATION, Ligne_Tableau_fils.FIL, Ligne_Tableau_fils.SECT, Ligne_Tableau_fils.TEINT,   "
Sql = Sql & "Ligne_Tableau_fils.TEINT2, Ligne_Tableau_fils.ISO , Ligne_Tableau_fils.Long, Ligne_Tableau_fils.[LONG CP],  "
Sql = Sql & " Ligne_Tableau_fils.Coupe, Ligne_Tableau_fils.POS, Ligne_Tableau_fils.[POS-OUT], Ligne_Tableau_fils.FA,   "
Sql = Sql & "Ligne_Tableau_fils.App, Ligne_Tableau_fils.VOI, Ligne_Tableau_fils.POS2, Ligne_Tableau_fils.[POS-OUT2],   "
Sql = Sql & "Ligne_Tableau_fils.FA2, Ligne_Tableau_fils.APP2, Ligne_Tableau_fils.VOI2, Ligne_Tableau_fils.PRECO, Ligne_Tableau_fils.Option  "
Sql = Sql & "FROM Ligne_Tableau_fils INNER JOIN Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart   "
Sql = Sql & "ON (Ligne_Tableau_fils.FIL = Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.FIL)   "
Sql = Sql & "AND (Ligne_Tableau_fils.Id_IndiceProjet = Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart.Id_IndiceProjet2)  "
Sql = Sql & "WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & TableProjet(0, i + 1) & " ;"

Debug.Print Sql

Set RsModifier = Con.OpenRecordSet(Sql)

L = 0: C = 0
If MajEcartExcel(Val(TableProjet(0, UBound(TableProjet, 2))), MyWorkbook, L, C, RsSuprimer, RsAjouter, RsModifier, "Tableau_fils_Ecart", txt) = True Then boolSave = True

Sql = "DROP TABLE Temp_" & NmJob & "_Connecteurs_Ecart;"
Con.Exequte Sql

Sql = " SELECT " & TableProjet(0, i + 1) & " AS Id_IndiceProjet2, Connecteurs.* INTO Temp_" & NmJob & "_Connecteurs_Ecart "
Sql = Sql & "FROM Connecteurs "
Sql = Sql & "WHERE Connecteurs.Id_IndiceProjet=" & TableProjet(0, i) & " ;"
Con.Exequte Sql


Sql = "SELECT Temp_" & NmJob & "_Connecteurs_Ecart.ACTIVER,Temp_" & NmJob & "_Connecteurs_Ecart.CONNECTEUR, Temp_" & NmJob & "_Connecteurs_Ecart.[O/N], Temp_" & NmJob & "_Connecteurs_Ecart.DESIGNATION, Temp_" & NmJob & "_Connecteurs_Ecart.CODE_APP,  "
Sql = Sql & "Temp_" & NmJob & "_Connecteurs_Ecart.N°, Temp_" & NmJob & "_Connecteurs_Ecart.POS, Temp_" & NmJob & "_Connecteurs_Ecart.[POS-OUT], Temp_" & NmJob & "_Connecteurs_Ecart.PRECO1,  "
Sql = Sql & "Temp_" & NmJob & "_Connecteurs_Ecart.PRECO2, Temp_" & NmJob & "_Connecteurs_Ecart.[100%], Temp_" & NmJob & "_Connecteurs_Ecart.OPTION "
Sql = Sql & "FROM Temp_" & NmJob & "_Connecteurs_Ecart LEFT JOIN Connecteurs ON (Temp_" & NmJob & "_Connecteurs_Ecart.N° = Connecteurs.N°)  "
Sql = Sql & "AND (Temp_" & NmJob & "_Connecteurs_Ecart.Id_IndiceProjet2 = Connecteurs.Id_IndiceProjet) "
Sql = Sql & "WHERE Temp_" & NmJob & "_Connecteurs_Ecart.Id_IndiceProjet=" & TableProjet(0, i + 1) & "  "
Sql = Sql & "AND Connecteurs.CODE_APP Is Null;"

Debug.Print Sql

Set RsSuprimer = Con.OpenRecordSet(Sql)

Sql = "SELECT Connecteurs.ACTIVER,Connecteurs.CONNECTEUR, Connecteurs.[O/N], Connecteurs.DESIGNATION, Connecteurs.CODE_APP,   "
Sql = Sql & "Connecteurs.N°, Connecteurs.POS, Connecteurs.[POS-OUT], Connecteurs.PRECO1, Connecteurs.PRECO2,   "
Sql = Sql & "Connecteurs.[100%], Connecteurs.OPTION  "
Sql = Sql & "FROM Connecteurs LEFT JOIN Temp_" & NmJob & "_Connecteurs_Ecart ON (Connecteurs.N° = Temp_" & NmJob & "_Connecteurs_Ecart.N°)   "
Sql = Sql & "AND (Connecteurs.Id_IndiceProjet = Temp_" & NmJob & "_Connecteurs_Ecart.Id_IndiceProjet2)  "
Sql = Sql & "WHERE Connecteurs.Id_IndiceProjet=" & TableProjet(0, i + 1) & " "
Sql = Sql & "AND Temp_" & NmJob & "_Connecteurs_Ecart.CODE_APP Is Null;"

Debug.Print Sql

Set RsAjouter = Con.OpenRecordSet(Sql)

Sql = "SELECT '' AS [Avant/Après],Temp_" & NmJob & "_Connecteurs_Ecart.ACTIVER, Temp_" & NmJob & "_Connecteurs_Ecart.CONNECTEUR, Temp_" & NmJob & "_Connecteurs_Ecart.[O/N], Temp_" & NmJob & "_Connecteurs_Ecart.DESIGNATION,  "
Sql = Sql & "Temp_" & NmJob & "_Connecteurs_Ecart.CODE_APP, Temp_" & NmJob & "_Connecteurs_Ecart.N°, Temp_" & NmJob & "_Connecteurs_Ecart.POS, Temp_" & NmJob & "_Connecteurs_Ecart.[POS-OUT],  "
Sql = Sql & "Temp_" & NmJob & "_Connecteurs_Ecart.PRECO1, Temp_" & NmJob & "_Connecteurs_Ecart.PRECO2, Temp_" & NmJob & "_Connecteurs_Ecart.[100%], Temp_" & NmJob & "_Connecteurs_Ecart.OPTION, '' AS Expr1,Connecteurs.ACTIVER,  "
Sql = Sql & "Connecteurs.CONNECTEUR, Connecteurs.[O/N], Connecteurs.DESIGNATION, Connecteurs.CODE_APP, Connecteurs.N°,  "
Sql = Sql & "Connecteurs.POS, Connecteurs.[POS-OUT], Connecteurs.PRECO1, Connecteurs.PRECO2, Connecteurs.[100%], Connecteurs.OPTION "
Sql = Sql & "FROM Temp_" & NmJob & "_Connecteurs_Ecart INNER JOIN Connecteurs ON (Temp_" & NmJob & "_Connecteurs_Ecart.N° = Connecteurs.N°)  "
Sql = Sql & "AND (Temp_" & NmJob & "_Connecteurs_Ecart.Id_IndiceProjet2 = Connecteurs.Id_IndiceProjet) "
Sql = Sql & "WHERE Connecteurs.Id_IndiceProjet=" & TableProjet(0, i + 1) & ";"
 

Debug.Print Sql

Set RsModifier = Con.OpenRecordSet(Sql)

L = 0: C = 0
If MajEcartExcel(Val(TableProjet(0, UBound(TableProjet, 2))), MyWorkbook, L, C, RsSuprimer, RsAjouter, RsModifier, "Connecteurs_Ecart", txt) = True Then boolSave = True

Sql = "DROP TABLE Temp_" & NmJob & "_T_Critères_Ecart;"
Con.Exequte Sql

Sql = "SELECT " & TableProjet(0, i + 1) & " AS Id_IndiceProjet2, T_Critères.* INTO Temp_" & NmJob & "_T_Critères_Ecart "
Sql = Sql & "FROM T_Critères "
Sql = Sql & "WHERE T_Critères.Id_IndiceProjet=" & TableProjet(0, i) & ";"

Con.Exequte Sql


Sql = "SELECT Temp_" & NmJob & "_T_Critères_Ecart.ACTIVER,Temp_" & NmJob & "_T_Critères_Ecart.CODE_CRITERE, Temp_" & NmJob & "_T_Critères_Ecart.CRITERES "
Sql = Sql & "FROM Temp_" & NmJob & "_T_Critères_Ecart LEFT JOIN T_Critères ON (Temp_" & NmJob & "_T_Critères_Ecart.CODE_CRITERE = T_Critères.CODE_CRITERE)  "
Sql = Sql & "AND (Temp_" & NmJob & "_T_Critères_Ecart.Id_IndiceProjet2 = T_Critères.Id_IndiceProjet) "
Sql = Sql & "WHERE Temp_" & NmJob & "_T_Critères_Ecart.Id_IndiceProjet=" & TableProjet(0, i + 1) & "  "
Sql = Sql & "AND T_Critères.CODE_CRITERE Is Null;"

Debug.Print Sql

Set RsSuprimer = Con.OpenRecordSet(Sql)

Sql = "SELECT T_Critères.ACTIVER,T_Critères.CODE_CRITERE, T_Critères.CRITERES "
Sql = Sql & "FROM T_Critères LEFT JOIN Temp_" & NmJob & "_T_Critères_Ecart ON (T_Critères.CODE_CRITERE = Temp_" & NmJob & "_T_Critères_Ecart.CODE_CRITERE)  "
Sql = Sql & "AND (T_Critères.Id_IndiceProjet = Temp_" & NmJob & "_T_Critères_Ecart.Id_IndiceProjet2) "
Sql = Sql & "WHERE T_Critères.Id_IndiceProjet=" & TableProjet(0, i + 1) & "  "
Sql = Sql & "AND Temp_" & NmJob & "_T_Critères_Ecart.CODE_CRITERE Is Null;"

Debug.Print Sql

Set RsAjouter = Con.OpenRecordSet(Sql)

Sql = "SELECT '' AS [Avant/Après],Temp_" & NmJob & "_T_Critères_Ecart.ACTIVER, Temp_" & NmJob & "_T_Critères_Ecart.CODE_CRITERE, Temp_" & NmJob & "_T_Critères_Ecart.CRITERES, '' AS Expr1, T_Critères.ACTIVER, "
Sql = Sql & "T_Critères.CODE_CRITERE, T_Critères.CRITERES "
Sql = Sql & "FROM Temp_" & NmJob & "_T_Critères_Ecart INNER JOIN T_Critères ON (Temp_" & NmJob & "_T_Critères_Ecart.CODE_CRITERE = T_Critères.CRITERES)  "
Sql = Sql & "AND (Temp_" & NmJob & "_T_Critères_Ecart.Id_IndiceProjet2 = T_Critères.Id_IndiceProjet) "
Sql = Sql & "Where T_Critères.Id_IndiceProjet = " & TableProjet(0, i + 1) & "  "
Sql = Sql & "ORDER BY T_Critères.CODE_CRITERE;"

Debug.Print Sql

Set RsModifier = Con.OpenRecordSet(Sql)

L = 0: C = 0
If MajEcartExcel(Val(TableProjet(0, UBound(TableProjet, 2))), MyWorkbook, L, C, RsSuprimer, RsAjouter, RsModifier, "Critères_Ecart", txt) = True Then boolSave = True
Next
Set RsAjouter = Con.CloseRecordSet(RsAjouter)
Set RsModifier = Con.CloseRecordSet(RsModifier)
Set RsSuprimer = Con.CloseRecordSet(RsSuprimer)


Sql = "DROP TABLE Temp_" & NmJob & "_Ecart_Nota;"
Con.Exequte Sql

Sql = "DROP TABLE Temp_" & NmJob & "_Composants_Ecart;"
Con.Exequte Sql

Sql = "DROP TABLE Temp_" & NmJob & "_T_Noeuds_Ecart;"
Con.Exequte Sql

Sql = "DROP TABLE Temp_" & NmJob & "_Ligne_Tableau_Fils_Ecart;"
Con.Exequte Sql

Sql = "DROP TABLE Temp_" & NmJob & "_Connecteurs_Ecart;"
Con.Exequte Sql

Sql = "DROP TABLE Temp_" & NmJob & "_T_Critères_Ecart;"
Con.Exequte Sql

If boolSave = True Then
Sql = "SELECT RqCartouche.* "
Sql = Sql & "FROM RqCartouche "
 Sql = Sql & "WHERE T_indiceProjet.Id=" & TableProjet(0, UBound(TableProjet, 2)) & ";"
Set RsModifier = Con.OpenRecordSet(Sql)
If RsModifier.EOF = False Then
        PathPl = PathArchive(PathArchiveAutocad, "" & RsModifier!Client, "" & RsModifier!CleAc, "" & RsModifier!Pieces, "Li", RsModifier.Fields("Li"), IdIndiceProjet, RsModifier.Fields("PI_Indice"), RsModifier.Fields("LI_Indice"), RsModifier!Version, True)
        PathPl = Replace(PathPl, RsModifier.Fields("Li") & "_" & Trim(RsModifier.Fields("LI_Indice")), "Ecart_Indice")
        Dim Fso As New FileSystemObject
        If Fso.FolderExists(PathPl) = False Then Fso.CreateFolder PathPl
        PathPl = PathPl & "\" & RsModifier.Fields("Li") & "_" & Trim(RsModifier.Fields("LI_Indice"))
RepriseSave:
        MyFormatDate = Format(Now, "yyyy-mm-dd-h-m-s")
        If Fso.FileExists(PathPl & "_Ecart_Indice_" & MyFormatDate & ".XLS") = True Then
        DoEvents
          GoTo RepriseSave
        End If
        MyWorkbook.SaveAs PathPl & "_Ecart_Indice_" & MyFormatDate
        
        If IdFils <> 0 Then
        Sql = "SELECT RqCartouche.* "
        Sql = Sql & "FROM RqCartouche "
        Sql = Sql & "WHERE T_indiceProjet.Id=" & IdFils & ";"
        Set RsModifier = Con.OpenRecordSet(Sql)
         PathPl2 = PathArchive(PathArchiveAutocad, "" & RsModifier!Client, "" & RsModifier!CleAc, "" & RsModifier!Pieces, "Li", RsModifier.Fields("Li"), IdFils, RsModifier.Fields("PI_Indice"), RsModifier.Fields("LI_Indice"), RsModifier!Version, True)
         
         PathPl2 = Replace(PathPl2, RsModifier.Fields("Li") & "_" & Trim(RsModifier.Fields("LI_Indice")), "Ecart_Indice")
         If Fso.FolderExists(PathPl2) = False Then Fso.CreateFolder PathPl2
        PathPl2 = PathPl2 & "\" & RsModifier.Fields("Li") & "_" & Trim(RsModifier.Fields("LI_Indice"))
        PathPl2 = PathPl2 & "_Ecart_Indice_" & MyFormatDate
       Racourci "" & PathPl2, "" & PathPl & "_Ecart_Indice_" & MyFormatDate, "XLS"
    End If
End If
End If

Sortie:

MyWorkbook.Close False
MyEcel.Quit
End Sub
