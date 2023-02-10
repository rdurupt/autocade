Attribute VB_Name = "ScanAutoCad"

Sub ScanDessin(Fichier As String, IdIndiceProjet As Long, Optional boolGarde As Boolean)
If boolAutoCAD = False Then Exit Sub
    Dim Fso As New FileSystemObject
    Dim NewBlock  As Object
    Dim NewBlock2  As Object
    Dim Entity As Object
    Dim BlocRef As Object
    Dim FicherSource As String
    Dim Attributes As Variant
   Dim Sql As String
Dim Rs As Recordset
Dim IdProjet As Long
Dim IdIndice As Long
DoEvents

Con.Execute "DELETE Xls_Nota.* FROM Xls_Nota WHERE Xls_Nota.Job=" & NmJob & ";"
Con.Execute "DELETE Xls_Connecteurs.* FROM Xls_Connecteurs WHERE Xls_Connecteurs.Job=" & NmJob & ";"
Con.Execute "DELETE Xls_Composants.* FROM Xls_Composants WHERE Xls_Composants.Job=" & NmJob & ";"
Con.Execute "DELETE xls_Ligne_Tableau_fils.* FROM xls_Ligne_Tableau_fils  where xls_Ligne_Tableau_fils.Job=" & NmJob & ";"
Con.Execute "DELETE Xls_Critères.* FROM Xls_Critères WHERE Xls_Critères.Job=" & NmJob & ";"
Con.Execute "DELETE Xls_Noeuds.* FROM Xls_Noeuds WHERE Xls_Noeuds.Job=" & NmJob & ";"

  
  Set TableauPath = funPath
    PathArchiveAutocad = TableauPath.Item("PathArchiveAutocad")
    PathArchiveAutocad = DefinirChemienComplet(TableauPath.Item("PathServer"), PathArchiveAutocad)
'     If Left(PathArchiveAutocad, 2) <> "\\" And Left(PathArchiveAutocad, 1) = "\" Then PathArchiveAutocad = TableauPath.Item("PathServer") & PathArchiveAutocad

 Sql = "SELECT T_indiceProjet.*, T_Pieces.Description as Pieces "
 Sql = Sql & "FROM T_Projet INNER JOIN (T_Pieces INNER JOIN  "
 Sql = Sql & "T_indiceProjet ON T_Pieces.Id = T_indiceProjet.Id_Pieces)  "
 Sql = Sql & "ON T_Projet.id = T_Pieces.IdProjet "
 Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"
    
    
    Set Rs = Con.OpenRecordSet(Sql)
If boolGarde = True And Rs.EOF = False Then
    FicherSource = Dir(Fichier)
    If FicherSource <> "" Then
     PathPl = PathArchive(TableauPath("PathArchiveAutocad"), "" & Rs!Client, "" & Rs!CleAc, "" & Rs!Pieces, Left(UCase(FicherSource), 2), Rs.Fields(Left(UCase(FicherSource), 2)), IdIndiceProjet, Rs.Fields("pi_Indice"), Rs.Fields(Left(UCase(FicherSource), 2) & "_Indice"), Rs!Version)
    End If
End If

    
'    'Set AutoApp = ThisDrawing.Application
    OpenFichier Fichier
'    AutoApp.Visible = True
     FormBarGrah.ProgressBar1Caption = " Scanne Tableau des Fils:"
     FormBarGrah.ProgressBar1.Value = 0
     FormBarGrah.ProgressBar1.Max = 1 + DocAutoCad.ModelSpace.Count
    
    For I = 0 To DocAutoCad.ModelSpace.Count - 1
     IncremanteBarGrah FormBarGrah
    DoEvents
        Set Entity = DocAutoCad.ModelSpace.Item(I)
       
        If Entity.ObjectName = "AcDbBlockReference" Then
            Set BlocRef = Entity
            a = BlocRef.Name
            If BlocRef.HasAttributes Then
                Attributes = BlocRef.GetAttributes
                 If (UBound(Attributes) = 15) Or (UBound(Attributes) = 14) Or (UBound(Attributes) = 13) Or (UBound(Attributes) = 12) Then
                    If IsTableauFils(Attributes) = True Then
                       EnrichirBaseFils Attributes
                    End If
                End If
            End If
        End If
    Next I

   Sql = "SELECT xls_Ligne_Tableau_fils.FIL FROM xls_Ligne_Tableau_fils "
    Sql = Sql & "Where xls_Ligne_Tableau_fils.Job =" & NmJob & " "
    Sql = Sql & "ORDER BY Val(xls_Ligne_Tableau_fils.FIL);"
    Set Rs = Con.OpenRecordSet(Sql)
    Dim IndexNum As Long
    IndexNum = 0
    While Rs.EOF = False
    IndexNum = IndexNum + 1
        If Val("" & Rs!Fil) > IndexNum Then
            For I = IndexNum To Val("" & Rs!Fil) - 1
                Sql = "INSERT INTO Xls_Connecteurs ( N°, Job, CONNECTEUR ) "
                Sql = Sql & "VALUES ('" & CStr(I) & "' , " & NmJob & " , 'ATTENTE' );"

                Con.Execute Sql

            Next
            IndexNum = I
        End If
        Rs.MoveNext
    Wend
    Set Rs = Con.CloseRecordSet(Rs)
       FormBarGrah.ProgressBar1Caption = " Scanne des Connecteurs:"
     FormBarGrah.ProgressBar1.Value = 0
     FormBarGrah.ProgressBar1.Max = 1 + DocAutoCad.ModelSpace.Count
    
    For I = 0 To DocAutoCad.ModelSpace.Count - 1
     IncremanteBarGrah FormBarGrah
    DoEvents
        Set Entity = DocAutoCad.ModelSpace.Item(I)
        If Entity.ObjectName = "AcDbBlockReference" Then
            Set BlocRef = Entity
            If BlocRef.HasAttributes Then
           a = BlocRef.Name
         
                Attributes = BlocRef.GetAttributes
                    If IsConnecteurs(Attributes) = True Then
                     If IsEpissures(Attributes) = True Then
                             EnrichirBaseConnecteurEpissure Attributes, BlocRef.Name
                     Else
                        
                         EnrichirBaseConnecteur Attributes, BlocRef.Name
                    
                      End If
                    End If
           End If
        End If
    Next I
    
    
    Sql = "SELECT Xls_Connecteurs.N° FROM Xls_Connecteurs "
    Sql = Sql & "Where Xls_Connecteurs.Job = " & NmJob & " "
    Sql = Sql & "ORDER BY Val(Xls_Connecteurs.N°);"
    Set Rs = Con.OpenRecordSet(Sql)
    IndexNum = 0
    
    While Rs.EOF = False
     IndexNum = IndexNum + 1
        If Val("" & Rs!N°) > IndexNum Then
            For I = IndexNum To Val("" & Rs!N°) - 1
                Sql = "INSERT  INTO Xls_Connecteurs ( N°, Job, CONNECTEUR )"
                Sql = Sql & "VALUES ( '" & CStr(I) & "' , " & NmJob & ",'ATTENTE');"
                Con.Execute Sql

            Next
            IndexNum = I
        End If
        Rs.MoveNext
    Wend
    
    

       FormBarGrah.ProgressBar1Caption = " Scanne des Composants:"
     FormBarGrah.ProgressBar1.Value = 0
     FormBarGrah.ProgressBar1.Max = 1 + DocAutoCad.ModelSpace.Count
    
    For I = 0 To DocAutoCad.ModelSpace.Count - 1
     IncremanteBarGrah FormBarGrah
    DoEvents
        Set Entity = DocAutoCad.ModelSpace.Item(I)
        If Entity.ObjectName = "AcDbBlockReference" Then
            Set BlocRef = Entity
            If BlocRef.HasAttributes Then
           a = BlocRef.Name
         
                Attributes = BlocRef.GetAttributes
                    If IsComposants(Attributes) = True Then
                           If UCase(BlocRef.Name) <> UCase("COMP_DESGN") Then
                         EnrichirBaseComposants Attributes, BlocRef.Name
                    End If
                      End If
                    
           End If
        End If
    Next I
    
   

     Sql = " SELECT Xls_Composants.NUMCOMP FROM Xls_Composants "
    Sql = Sql & "Where Xls_Composants.Job= " & NmJob & " "
    Sql = Sql & "ORDER BY Val(Xls_Composants.NUMCOMP);"
    Set Rs = Con.OpenRecordSet(Sql)
    IndexNum = 0
    
    While Rs.EOF = False
     IndexNum = IndexNum + 1
        If Val("" & Rs!NUMCOMP) > IndexNum Then
            For I = IndexNum To Val("" & Rs!NUMCOMP) - 1
                Sql = "INSERT  INTO Xls_Composants ( NUMCOMP, Job, DESIGNCOMP )"
                Sql = Sql & "VALUES ( '" & CStr(I) & "' , " & NmJob & ",'ATTENTE');"
                Con.Execute Sql

            Next
            IndexNum = I
        End If
        Rs.MoveNext
    Wend

       FormBarGrah.ProgressBar1Caption = " Scanne des Notas:"
     FormBarGrah.ProgressBar1.Value = 0
     FormBarGrah.ProgressBar1.Max = 1 + DocAutoCad.ModelSpace.Count
    
    For I = 0 To DocAutoCad.ModelSpace.Count - 1
     IncremanteBarGrah FormBarGrah
    DoEvents
        Set Entity = DocAutoCad.ModelSpace.Item(I)
        If Entity.ObjectName = "AcDbBlockReference" Then
            Set BlocRef = Entity
            If BlocRef.HasAttributes Then
           a = BlocRef.Name
         
                Attributes = BlocRef.GetAttributes
                    If IsNotas(Attributes) = True Then
                                          
                         EnrichirBaseNota Attributes, BlocRef.Name
                    
                      End If
                    
           End If
        End If
    Next I
    
    

     Sql = "SELECT Xls_Nota.NUMNOTA FROM Xls_Nota "
    Sql = Sql & "Where Xls_Nota.Job= " & NmJob & " "
    Sql = Sql & "ORDER BY Val(Xls_Nota.NUMNOTA);"
    Set Rs = Con.OpenRecordSet(Sql)
    IndexNum = 0
    
    While Rs.EOF = False
     IndexNum = IndexNum + 1
        If Val("" & Rs!NUMNOTA) > IndexNum Then
            For I = IndexNum To Val("" & Rs!NUMNOTA) - 1
                Sql = "INSERT  INTO Xls_Nota ( NUMNOTA, Job, NOTA )"
                Sql = Sql & "VALUES ( '" & CStr(I) & "' , " & NmJob & ",'ATTENTE');"
                Con.Execute Sql

            Next
            IndexNum = I
        End If
        Rs.MoveNext
    Wend

  FormBarGrah.ProgressBar1Caption = " Scanne des Noeds:"
     FormBarGrah.ProgressBar1.Value = 0
     FormBarGrah.ProgressBar1.Max = 1 + DocAutoCad.ModelSpace.Count
    
    For I = 0 To DocAutoCad.ModelSpace.Count - 1
     IncremanteBarGrah FormBarGrah
    DoEvents
        Set Entity = DocAutoCad.ModelSpace.Item(I)
        If Entity.ObjectName = "AcDbBlockReference" Then
            Set BlocRef = Entity
            If BlocRef.HasAttributes Then
           a = BlocRef.Name
         
                Attributes = BlocRef.GetAttributes
                    If IsNoeuds(Attributes) = True Then
                          
                         EnrichirBaseNoeuds Attributes, BlocRef.Name
                    
                      End If
                    
           End If
        End If
    Next I
    
Sql = "SELECT T_Regle_Comp_Hab.ENCELADE  FROM T_Regle_Comp_Hab "
Sql = Sql & "WHERE T_Regle_Comp_Hab.ENCELADE Is Not Null "
Sql = Sql & "And T_Regle_Comp_Hab.ENCELADE<>'' "
Sql = Sql & "ORDER BY T_Regle_Comp_Hab.ENCELADE;"
Set Rs = Con.OpenRecordSet(Sql)
Dim ValNoeud As String
    If Rs.EOF = False Then
        ValNoeud = "" & Rs!ENCELADE
    End If
 Sql = "SELECT Xls_Noeuds.NŒUDS FROM Xls_Noeuds "
Sql = Sql & "Where Xls_Noeuds.Job = " & NmJob & " "
Sql = Sql & "ORDER BY Xls_Noeuds.NŒUDS;"
Set Rs = Con.OpenRecordSet(Sql)
Dim IndexCodeNeouds As Long
IndexCodeNeouds = 2
While Rs.EOF = False
aa = NoeuName(IndexCodeNeouds)
    While Trim(UCase("" & Rs!NŒUDS)) <> aa
        Sql = "INSERT INTO Xls_Noeuds ( NŒUDS, Job,CODE_ENC) "
        Sql = Sql & "VALUES ( '" & aa & "' , " & NmJob & ",'" & MyReplace(ValNoeud) & "') ;"
        Con.Execute Sql
        IndexCodeNeouds = IndexCodeNeouds + 1
        aa = NoeuName(IndexCodeNeouds)
    Wend
    IndexCodeNeouds = IndexCodeNeouds + 1
    Rs.MoveNext
Wend
    FormBarGrah.ProgressBar1Caption = " Scanne des Critères:"
     FormBarGrah.ProgressBar1.Value = 0
     FormBarGrah.ProgressBar1.Max = 1 + DocAutoCad.ModelSpace.Count
    
    For I = 0 To DocAutoCad.ModelSpace.Count - 1
     IncremanteBarGrah FormBarGrah
    DoEvents
        Set Entity = DocAutoCad.ModelSpace.Item(I)
        If Entity.ObjectName = "AcDbBlockReference" Then
            Set BlocRef = Entity
            If BlocRef.HasAttributes Then
           a = BlocRef.Name
         
                Attributes = BlocRef.GetAttributes
                    If IsCriteres(Attributes) = True Then
                         
                         EnrichirBaseCritères Attributes, BlocRef.Name
                   
                      End If
                    
           End If
        End If
    Next I
    MajBase IdIndiceProjet
   
    
     FormBarGrah.ProgressBar1Caption = " Traitement terminé:"
     FormBarGrah.ProgressBar1.Value = 0
     If Trim("" & PathPl) <> "" Then SaveAs PathPl
   CloseDocument
    End Sub
    
Function ReplaceAttribs(Txt As String, Sql As String) As String
Dim boolReplace As Boolean
ReplaceAttribs = Txt
If UCase(Txt) = "CO" Then
    If InStr(1, Sql, "TEINT") <> 0 Then
        If Len(Txt) = 2 Then ReplaceAttribs = Replace(UCase(ReplaceAttribs), "CO", "TEINT2")
    Else
        If Len(Txt) = 2 Then ReplaceAttribs = Replace(UCase(ReplaceAttribs), "CO", "TEINT")

    End If
End If

If UCase(Txt) = "CON" Then
    If InStr(1, Sql, "FA") <> 0 Then
        If Len(Txt) = 3 Then ReplaceAttribs = Replace(UCase(ReplaceAttribs), "CON", "FA2")
    Else
        If Len(Txt) = 3 Then ReplaceAttribs = Replace(UCase(ReplaceAttribs), "CON", "FA")

    End If
End If
If UCase(Txt) = "VOIE" Then
    If InStr(1, Sql, "VOI") <> 0 Then
        If Len(Txt) = 4 Then ReplaceAttribs = Replace(UCase(ReplaceAttribs), "VOIE", "VOI")
    Else
        If Len(Txt) = 4 Then ReplaceAttribs = Replace(UCase(ReplaceAttribs), "VOIE", "VOI2")

    End If
End If

If InStr(1, Sql, "POS") <> 0 Then
    If Len(Txt) = 3 Then ReplaceAttribs = Replace(UCase(ReplaceAttribs), "POS", "POS2")
End If
ReplaceAttribs = "[" & ReplaceAttribs & "]"
End Function
Sub EnrichirBaseFils(Attributes As Variant)
    Dim Sql As String
    Dim SqlValues As String
    Dim sqlNull As Boolean
    Dim Rs As Recordset
    
    sqlNull = True
    Num = 0
    Sql = "INSERT INTO xls_Ligne_Tableau_fils ( Job,ACTIVER,"
    For I = LBound(Attributes) To UBound(Attributes)
        Debug.Print ReplaceAttribs(Attributes(I).TagString, Sql)
        Sql = Sql & ReplaceAttribs(Attributes(I).TagString, Sql) & ","
    Next I
    Sql = Left(Sql, Len(Sql) - 1)
    Sql = Sql & ") Values ("
    SqlValues = ""
    
    SqlValues = SqlValues & NmJob & ",true,"
    For I = LBound(Attributes) To UBound(Attributes)
    If UCase(Trim(Attributes(I).TextString)) = "FIL" Then Exit Sub
        If Trim("" & Attributes(I).TextString) = "" Then
            SqlValues = SqlValues & "NULL,"
        Else
            sqlNull = False
            SqlValues = SqlValues & "'" & MyReplace("" & Attributes(I).TextString) & "',"
        End If
    Next I
    
    SqlValues = Left(SqlValues, Len(SqlValues) - 1)
      Sql = Sql & SqlValues & ");"
      If sqlNull = False Then Con.Execute Sql
End Sub
Sub EnrichirBaseConnecteur(Attributes As Variant, NameConnecteur)
    Dim Sql As String
    Dim SqlValues As String
    Dim Rs As Recordset
    Dim Table(5) As String
    
    Static Ip As Long
    Ip = Ip + 1
Table(0) = "DESIGNATION"
Table(1) = "POS"
Table(2) = "N°"
Table(3) = "CODE_APP"
Table(4) = "PRECO1"
Table(5) = "PRECO2"
    
    Sql = "INSERT INTO Xls_Connecteurs (Job,ACTIVER,CONNECTEUR, [O/N], DESIGNATION,POS, N°,   CODE_APP, PRECO1, PRECO2 )"
    Sql = Sql & "Values ("
    SqlValues = ""
    
    SqlValues = SqlValues & NmJob & ",true,'" & MyReplace("" & NameConnecteur) & "',false,"
    For I = LBound(Table) To UBound(Table)
        For I2 = LBound(Attributes) To UBound(Attributes)
            If Table(I) = PRECO(UCase("" & Attributes(I2).TagString)) Then
            
                If Trim("" & Attributes(I2).TextString) = "" Then
                    SqlValues = SqlValues & "NULL,"
                Else
                    SqlValues = SqlValues & "'" & MyReplace("" & Attributes(I2).TextString) & "',"
                End If
                Exit For
            End If
       
        Next I2
    Next I
    SqlValues = Left(SqlValues, Len(SqlValues) - 1)
    Sql = Sql & SqlValues & ");"
    Debug.Print Sql
    Con.Execute Sql
End Sub

Sub EnrichirBaseComposants(Attributes As Variant, NameConnecteur)
    Dim Sql As String
    Dim SqlValues As String
    Dim Rs As Recordset
    Dim Table(3) As String
    
    Static Ip As Long
    Ip = Ip + 1
Table(0) = "DESIGNCOMP"
Table(1) = "NUMCOMP"
Table(2) = "REFCOMP"
Table(3) = "PATHCOMP"
    
    Sql = "INSERT INTO Xls_Composants (Job,ACTIVER,DESIGNCOMP, NUMCOMP, REFCOMP,Path )"
    Sql = Sql & "Values ("
    SqlValues = ""
    
    SqlValues = SqlValues & NmJob & ",true,"
    For I = LBound(Table) To UBound(Table)
        For I2 = LBound(Attributes) To UBound(Attributes)
            If Table(I) = PRECO(UCase("" & Attributes(I2).TagString)) Then
            
                If Trim("" & Attributes(I2).TextString) = "" Then
                    SqlValues = SqlValues & "NULL,"
                Else
                    If Attributes(I2).TagString = "NUMCOMP" Then
                    
                        SqlValues = SqlValues & "" & CInt(Mid(Attributes(I2).TextString, 2, Len(Attributes(I2).TextString) - 1)) & ","
                    Else
                        SqlValues = SqlValues & "'" & MyReplace("" & Attributes(I2).TextString) & "',"
                    End If
                End If
                Exit For
            End If
       
        Next I2
    Next I
    SqlValues = Left(SqlValues, Len(SqlValues) - 1)
    Sql = Sql & SqlValues & ");"
    Debug.Print Sql
    Con.Execute Sql
End Sub
Sub EnrichirBaseNoeuds(Attributes As Variant, NameConnecteur)
    Dim Sql As String
    Dim SqlValues As String
    Dim Rs As Recordset
    Dim Table(4) As String
    Dim Trouve As Boolean
    Table(0) = UCase("NOEUD")
    Table(1) = UCase("LONG")
    Table(2) = UCase("HAB")
    Table(3) = UCase("DIAM")
    Table(4) = UCase("CLASSE_T")
'    Table(5) = UCase("LONGUEUR_CUMULEE")
   
    
    Static Ip As Long
    Ip = Ip + 1

'    Dim Trouve As Boolean
    Set Colec = ColectionAttribueConecteur(Attributes)
Sql = "SELECT Xls_Noeuds.NŒUDS, Xls_Noeuds.Job "
Sql = Sql & "FROM Xls_Noeuds "
Sql = Sql & "WHERE Xls_Noeuds.NŒUDS='" & Attributes(Colec("NOEUD")).TextString & "' "
Sql = Sql & "AND Xls_Noeuds.Job=" & NmJob & ";"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = True Then
    Sql = "INSERT INTO Xls_Noeuds (Job ,ACTIVER, NŒUDS,"
    Sql = Sql & "LONGUEUR, CODE_ENC, DIAMETRE, CLASSE_T  "
    Sql = Sql & " ) "
    Sql = Sql & "Values ("
    SqlValues = ""
    
    SqlValues = SqlValues & NmJob & ",'1',"
    For I = LBound(Table) To UBound(Table)
    Trouve = False
        For I2 = LBound(Attributes) To UBound(Attributes)
            If Table(I) = PRECO(UCase("" & Attributes(I2).TagString)) Then
            Trouve = True
                If Trim("" & Attributes(I2).TextString) = "" Then
                    SqlValues = SqlValues & "NULL,"
                Else
                
                    SqlValues = SqlValues & "'" & MyReplace("" & Attributes(I2).TextString) & "',"
                End If
                Exit For
            End If
       
        Next I2
        If Trouve = False Then
            SqlValues = SqlValues & "NULL,"
        End If
    Next I
    SqlValues = Left(SqlValues, Len(SqlValues) - 1)
    Sql = Sql & SqlValues & ");"
   
    Debug.Print Sql
    Con.Execute Sql
    End If
    Sql = "UPDATE Xls_Noeuds SET "
    For I = 0 To Colec.Count - 1
     Sql = Sql & "Xls_Noeuds." & Replace(Replace(Replace(Replace(Replace(PRECO(UCase("" & Attributes(I).TagString)), "HAB", "CODE_ENC"), "LONG", "LONGUEUR"), "DIAM", "DIAMETRE"), "NOEUD", "NŒUDS"), "LONGUEUR_CUMUL", "LONGUEUR_CUMULEE") & " = '" & MyReplace(Attributes(I).TextString) & "',"
    Next
    Sql = Left(Sql, Len(Sql) - 1)
Sql = Sql & " WHERE Xls_Noeuds.NŒUDS='" & Attributes(Colec("NOEUD")).TextString & "' "
Sql = Sql & "AND Xls_Noeuds.Job=" & NmJob & ";"
Con.Execute Sql
Sql = ""
Select Case UCase(NameConnecteur)
    Case "NOEUD_PRINCIPAL"
            Sql = "UPDATE Xls_Noeuds SET Xls_Noeuds.TORON_PRINCIPAL = 1, Xls_Noeuds.Fleche_Droite = '0' "
         
    Case "NOEUD_PRINCIPAL1"
            Sql = "UPDATE Xls_Noeuds SET Xls_Noeuds.TORON_PRINCIPAL = 1, Xls_Noeuds.Fleche_Droite = '1' "
            
             
    Case "NOEUD_SECONDAIRE"
            Sql = "UPDATE Xls_Noeuds SET Xls_Noeuds.TORON_PRINCIPAL = 0, Xls_Noeuds.Fleche_Droite = '0'"
     Case "NOEUD_SECONDAIRE1"
            Sql = "UPDATE Xls_Noeuds SET Xls_Noeuds.TORON_PRINCIPAL = 0, Xls_Noeuds.Fleche_Droite = '1' "
End Select
If Sql <> "" Then
    Sql = Sql & " WHERE Xls_Noeuds.NŒUDS='" & Attributes(Colec("NOEUD")).TextString & "' "
    Sql = Sql & "AND Xls_Noeuds.Job=" & NmJob & ";"
    Con.Execute Sql
End If
End Sub
Sub EnrichirBaseCritères(Attributes As Variant, NameConnecteur)
    Dim Sql As String
    Dim SqlValues As String
    Dim Rs As Recordset
    Dim Table(1) As String
    Dim Trouve As Boolean
    Table(0) = UCase("REFCRITERE")
    Table(1) = UCase("REFCRITERELIB")
   
   
    
    Static Ip As Long
    Ip = Ip + 1

'    Dim Trouve As Boolean
    

    Sql = "INSERT INTO Xls_Critères (Job,ACTIVER,CODE_CRITERE, CRITERES)"
    Sql = Sql & "Values (" & NmJob & ",true,"
    SqlValues = ""
    
'    SqlValues = SqlValues & NmJob & ",'" & MyReplace("" & NameConnecteur) & "',"
    For I = LBound(Table) To UBound(Table)
    Trouve = False
        For I2 = LBound(Attributes) To UBound(Attributes)
            If Table(I) = PRECO(UCase("" & Attributes(I2).TagString)) Then
            Trouve = True
            If Trim("" & Attributes(I2).TextString) = "CRITERE" Then
                Exit Sub
            End If
            
                If Trim("" & Attributes(I2).TextString) = "" Then
                    SqlValues = SqlValues & "NULL,"
                Else
                
                    SqlValues = SqlValues & "'" & MyReplace("" & Attributes(I2).TextString) & "',"
                End If
                Exit For
            End If
       
        Next I2
        If Trouve = False Then
            SqlValues = SqlValues & "NULL,"
        End If
    Next I
    SqlValues = Left(SqlValues, Len(SqlValues) - 1)
    Sql = Sql & SqlValues & ");"
    Debug.Print Sql
    Con.Execute Sql
End Sub
Sub EnrichirBaseNota(Attributes As Variant, NameConnecteur)
    Dim Sql As String
    Dim SqlValues As String
    Dim Rs As Recordset
    Dim Table(0) As String
    
    Static Ip As Long
    Ip = Ip + 1

    Dim Trouve As Boolean
    
Table(0) = "NUMNOTA"
    Sql = "INSERT INTO Xls_Nota (Job,ACTIVER,NOTA, NUMNOTA )"
    Sql = Sql & "Values ("
    SqlValues = ""
    
    SqlValues = SqlValues & NmJob & ",true,'" & NameConnecteur & "',"
    For I = LBound(Table) To UBound(Table)
        For I2 = LBound(Attributes) To UBound(Attributes)
            If Table(I) = PRECO(UCase("" & Attributes(I2).TagString)) Then
            
                If Trim("" & Attributes(I2).TextString) = "" Then
                    SqlValues = SqlValues & "NULL,"
                Else
                
                    SqlValues = SqlValues & "'" & MyReplace("" & Attributes(I2).TextString) & "',"
                End If
                Exit For
            End If
       
        Next I2
    Next I
    SqlValues = Left(SqlValues, Len(SqlValues) - 1)
    Sql = Sql & SqlValues & ");"
    Debug.Print Sql
    Con.Execute Sql
End Sub
Sub EnrichirBaseConnecteurEpissure(Attributes As Variant, NameConnecteur)
    Dim Sql As String
    Dim SqlValues As String
    Dim Rs As Recordset
    Dim Table(5) As String
   
  Table(0) = "DESIGNATION"

Table(1) = "N°"
Table(2) = "CODE_APP"
Table(3) = "POS"
Table(4) = "PRECO1"
Table(5) = "PRECO2"
    
    Sql = "INSERT INTO Xls_Connecteurs (Job,ACTIVER, CONNECTEUR, [O/N], DESIGNATION,  N°, CODE_APP,POS, PRECO1, PRECO2 )"
    Sql = Sql & "Values ("
    SqlValues = ""
        SqlValues = ""
        
    SqlValues = SqlValues & NmJob & ",true,'" & MyReplace("" & NameConnecteur) & "',true,"

     For I = LBound(Table) To UBound(Table)
        For I2 = LBound(Attributes) To UBound(Attributes)
            If Table(I) = PRECO(UCase("" & Attributes(I2).TagString)) Then
            
                If Trim("" & Attributes(I2).TextString) = "" Then
                    SqlValues = SqlValues & "NULL,"
                Else
                    SqlValues = SqlValues & "'" & MyReplace("" & Attributes(I2).TextString) & "',"
                End If
                Exit For
            End If
       
        Next I2
    Next I
    SqlValues = Left(SqlValues, Len(SqlValues) - 1)
    Sql = Sql & SqlValues & ");"
    Debug.Print Sql
    Con.Execute Sql
End Sub
Sub MajImportAutocad(Id_IndiceProjet)
Dim Sql As String
Dim SqlValues As String
Dim IConnecteur As Long
Dim boolReprise As Boolean
Dim Rs As Recordset
Sql = "DELETE Connecteurs.*, Connecteurs.Id_IndiceProjet "
Sql = Sql & "FROM Connecteurs "
Sql = Sql & "WHERE Connecteurs.Id_IndiceProjet=" & Id_IndiceProjet & ";"

Con.Execute Sql

Sql = "DELETE Ligne_Tableau_fils.*, Ligne_Tableau_fils.Id_IndiceProjet "
Sql = Sql & "FROM Ligne_Tableau_fils "
Sql = Sql & "WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & Id_IndiceProjet & ";"
Con.Execute Sql

Sql = "SELECT " & Id_IndiceProjet & " AS Id_IndiceProjet, RqXls_Connecteurs.CONNECTEUR, RqXls_Connecteurs.[O/N], RqXls_Connecteurs.DESIGNATION, RqXls_Connecteurs.CODE_APP, RqXls_Connecteurs.N°, RqXls_Connecteurs.POS, RqXls_Connecteurs.PRECO1, RqXls_Connecteurs.PRECO2 "
Sql = Sql & "FROM RqXls_Connecteurs "
Sql = Sql & "WHERE RqXls_Connecteurs.Job=" & NmJob & ";"
Set Rs = Con.OpenRecordSet(Sql)
IConnecteur = 0
While Rs.EOF = False

Sql = "INSERT INTO Connecteurs ( Id_IndiceProjet, CONNECTEUR, [O/N], DESIGNATION, CODE_APP, N°, POS, PRECO1, PRECO2 ) "
Sql = Sql & "VALUES("
Reprise:
IConnecteur = IConnecteur + 1

SqlValues = Id_IndiceProjet & ","
If IConnecteur < Val("" & Rs![N°]) Then
    boolReprise = True
     SqlValues = SqlValues & "'NEANT',"
    For I = 2 To Rs.Fields.Count - 1
            If Rs.Fields(I).Name = "N°" Then
                SqlValues = SqlValues & "'" & IConnecteur & "',"
            Else
                SqlValues = SqlValues & "NULL,"
            End If
    Next I

Else
    boolReprise = False
For I = 1 To Rs.Fields.Count - 1
    If Trim("" & Rs.Fields(I)) = "" Then
        SqlValues = SqlValues & "NULL,"
    Else
        If Rs.Fields(I).Type = adBoolean Then
             SqlValues = SqlValues & Replace(Replace(Rs.Fields(I), "Faux", "false"), "Vrai", "true") & " ,"
        Else
        SqlValues = SqlValues & "'" & Trim("" & Rs.Fields(I)) & "',"
        End If
    End If
Next I
End If
    SqlValues = Left(SqlValues, Len(SqlValues) - 1)
    Con.Execute Sql & SqlValues & ");"
    If boolReprise = True Then GoTo Reprise
Rs.MoveNext
Wend

Sql = "SELECT  Connecteurs.N° "
Sql = Sql & "FROM Connecteurs "
Sql = Sql & "WHERE Connecteurs.Id_IndiceProjet = " & Id_IndiceProjet & " "
Sql = Sql & "And Connecteurs.N° Is Not Null "
Sql = Sql & "ORDER BY Connecteurs.N° DESC;"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = False Then
NbCon = Rs!N°
Sql = "SELECT  Connecteurs.Numéro "
Sql = Sql & "FROM Connecteurs "
Sql = Sql & "WHERE Connecteurs.Id_IndiceProjet = " & Id_IndiceProjet & " "
Sql = Sql & "And Connecteurs.N° Is Null "

Set Rs = Con.OpenRecordSet(Sql)
While Rs.EOF = False
NbCon = NbCon + 1
Sql = "UPDATE Connecteurs SET Connecteurs.N° = " & NbCon & " "
Sql = Sql & "WHERE Connecteurs.Numéro=" & Rs!Numéro & ";"
Con.Execute Sql
    Rs.MoveNext
Wend
End If



Set Rs = Con.CloseRecordSet(Rs)
Sql = "INSERT INTO Ligne_Tableau_fils ( Id_IndiceProjet, LIAI, "
Sql = Sql & "DESIGNATION, FIL, SECT, TEINT, TEINT2, ISO, POS, "
Sql = Sql & "FA, VOI, POS2, FA2, VOI2, [LONG],APP,APP2 ) "
Sql = Sql & "SELECT " & Id_IndiceProjet & " AS Id_IndiceProjet, xls_Ligne_Tableau_fils.LIAI, "
Sql = Sql & "xls_Ligne_Tableau_fils.DESIGNATION, xls_Ligne_Tableau_fils.FIL, "
Sql = Sql & "xls_Ligne_Tableau_fils.SECT, xls_Ligne_Tableau_fils.TEINT, "
Sql = Sql & "xls_Ligne_Tableau_fils.TEINT2, xls_Ligne_Tableau_fils.ISO, "
Sql = Sql & "xls_Ligne_Tableau_fils.POS, xls_Ligne_Tableau_fils.FA, "
Sql = Sql & "xls_Ligne_Tableau_fils.VOI, xls_Ligne_Tableau_fils.POS2, "
Sql = Sql & "xls_Ligne_Tableau_fils.FA2, xls_Ligne_Tableau_fils.VOI2, "
Sql = Sql & "xls_Ligne_Tableau_fils.LONG, "
Sql = Sql & "(SELECT  Xls_Connecteurs.CODE_APP FROM Xls_Connecteurs WHERE Xls_Connecteurs.N°=[FA] and Xls_Connecteurs.Job=" & NmJob & ") as APP, "
Sql = Sql & "(SELECT  Xls_Connecteurs.CODE_APP FROM Xls_Connecteurs WHERE Xls_Connecteurs.N°=[FA2] and Xls_Connecteurs.Job=" & NmJob & ") as APP2 "

Sql = Sql & "FROM xls_Ligne_Tableau_fils "
Sql = Sql & "where xls_Ligne_Tableau_fils.Job=" & NmJob & ";"
Con.Execute Sql
End Sub
