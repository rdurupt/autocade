Attribute VB_Name = "ScanAutoCad"

Public Sub ScanDessin(Fichier As String, IdIndiceProjet As Long, Optional boolGarde As Boolean)
    Dim Fso As New FileSystemObject
    Dim NewBlock  As AcadBlock
    Dim NewBlock2  As AcadBlockReference
    Dim Entity As AcadEntity
    Dim BlocRef As AcadBlockReference
    Dim FicherSource As String
    Dim Attributes As Variant
   Dim sql As String
Dim Rs As Recordset
Dim IdProjet As Long
Dim IdIndice As Long
DoEvents

Con.Exequte "DELETE Xls_Nota.* FROM Xls_Nota WHERE Xls_Nota.Job=" & NmJob & ";"
Con.Exequte "DELETE Xls_Connecteurs.* FROM Xls_Connecteurs WHERE Xls_Connecteurs.Job=" & NmJob & ";"
Con.Exequte "DELETE Xls_Composants.* FROM Xls_Composants WHERE Xls_Composants.Job=" & NmJob & ";"
Con.Exequte "DELETE xls_Ligne_Tableau_fils.* FROM xls_Ligne_Tableau_fils  where xls_Ligne_Tableau_fils.Job=" & NmJob & ";"
 
  
  Set TableauPath = funPath
    PathArchiveAutocad = TableauPath.Item("PathArchiveAutocad")
     If Left(PathArchiveAutocad, 2) <> "\\" Then PathArchiveAutocad = TableauPath.Item("PathServer") & PathArchiveAutocad

 sql = "SELECT T_indiceProjet.*, T_Pieces.Description as Pieces "
 sql = sql & "FROM T_Projet INNER JOIN (T_Pieces INNER JOIN  "
 sql = sql & "T_indiceProjet ON T_Pieces.Id = T_indiceProjet.Id_Pieces)  "
 sql = sql & "ON T_Projet.id = T_Pieces.IdProjet "
 sql = sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"
    
    
    Set Rs = Con.OpenRecordSet(sql)
If boolGarde = True And Rs.EOF = False Then
    FicherSource = Dir(Fichier)
    If FicherSource <> "" Then
     PathPl = PathArchive(PathArchiveAutocad, "" & Rs!Client, "" & Rs!CleAc, "" & Rs!Pieces, Left(UCase(FicherSource), 2), Rs.Fields(Left(UCase(FicherSource), 2)), IdIndiceProjet, Rs.Fields("pi_Indice"), Rs.Fields(Left(UCase(FicherSource), 2) & "_Indice"), Rs!Version)
    End If
End If

    
'    Set AutoApp = ThisDrawing.Application
    OpenFichier Fichier
     FormBarGrah.ProgressBar1Caption = "Lecture Tableau des Fils:"
     FormBarGrah.ProgressBar1.Value = 0
     FormBarGrah.ProgressBar1.Max = 1 + AutoApp.ActiveDocument.ModelSpace.Count
    
    For i = 0 To AutoApp.ActiveDocument.ModelSpace.Count - 1
     FormBarGrah.ProgressBar1.Value = FormBarGrah.ProgressBar1.Value + 1
    DoEvents
        Set Entity = AutoApp.ActiveDocument.ModelSpace.Item(i)
       
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
    Next i

   

       FormBarGrah.ProgressBar1Caption = "Lecture des Connecteurs:"
     FormBarGrah.ProgressBar1.Value = 0
     FormBarGrah.ProgressBar1.Max = 1 + AutoApp.ActiveDocument.ModelSpace.Count
    
    For i = 0 To AutoApp.ActiveDocument.ModelSpace.Count - 1
     FormBarGrah.ProgressBar1.Value = FormBarGrah.ProgressBar1.Value + 1
    DoEvents
        Set Entity = AutoApp.ActiveDocument.ModelSpace.Item(i)
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
    Next i
    
    
    
    
    

       FormBarGrah.ProgressBar1Caption = "Lecture des Composants:"
     FormBarGrah.ProgressBar1.Value = 0
     FormBarGrah.ProgressBar1.Max = 1 + AutoApp.ActiveDocument.ModelSpace.Count
    
    For i = 0 To AutoApp.ActiveDocument.ModelSpace.Count - 1
     FormBarGrah.ProgressBar1.Value = FormBarGrah.ProgressBar1.Value + 1
    DoEvents
        Set Entity = AutoApp.ActiveDocument.ModelSpace.Item(i)
        If Entity.ObjectName = "AcDbBlockReference" Then
            Set BlocRef = Entity
            If BlocRef.HasAttributes Then
           a = BlocRef.Name
         
                Attributes = BlocRef.GetAttributes
                    If IsComposants(Attributes) = True Then
                                          
                         EnrichirBaseComposants Attributes, BlocRef.Name
                    
                      End If
                    
           End If
        End If
    Next i
    
   

       FormBarGrah.ProgressBar1Caption = "Lecture des Notas:"
     FormBarGrah.ProgressBar1.Value = 0
     FormBarGrah.ProgressBar1.Max = 1 + AutoApp.ActiveDocument.ModelSpace.Count
    
    For i = 0 To AutoApp.ActiveDocument.ModelSpace.Count - 1
     FormBarGrah.ProgressBar1.Value = FormBarGrah.ProgressBar1.Value + 1
    DoEvents
        Set Entity = AutoApp.ActiveDocument.ModelSpace.Item(i)
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
    Next i
    
    MajBase IdIndiceProjet
   
    
     FormBarGrah.ProgressBar1Caption = "Traitement terminé:"
     FormBarGrah.ProgressBar1.Value = 0
      SaveAs PathPl
   CloseDocument
    End Sub
    
Public Function ReplaceAttribs(txt As String, sql As String) As String
Dim boolReplace As Boolean
ReplaceAttribs = txt
If UCase(txt) = "CO" Then
    If InStr(1, sql, "TEINT") <> 0 Then
        If Len(txt) = 2 Then ReplaceAttribs = Replace(UCase(ReplaceAttribs), "CO", "TEINT2")
    Else
        If Len(txt) = 2 Then ReplaceAttribs = Replace(UCase(ReplaceAttribs), "CO", "TEINT")

    End If
End If

If UCase(txt) = "CON" Then
    If InStr(1, sql, "FA") <> 0 Then
        If Len(txt) = 3 Then ReplaceAttribs = Replace(UCase(ReplaceAttribs), "CON", "FA2")
    Else
        If Len(txt) = 3 Then ReplaceAttribs = Replace(UCase(ReplaceAttribs), "CON", "FA")

    End If
End If
If UCase(txt) = "VOIE" Then
    If InStr(1, sql, "VOI") <> 0 Then
        If Len(txt) = 4 Then ReplaceAttribs = Replace(UCase(ReplaceAttribs), "VOIE", "VOI")
    Else
        If Len(txt) = 4 Then ReplaceAttribs = Replace(UCase(ReplaceAttribs), "VOIE", "VOI2")

    End If
End If

If InStr(1, sql, "POS") <> 0 Then
    If Len(txt) = 3 Then ReplaceAttribs = Replace(UCase(ReplaceAttribs), "POS", "POS2")
End If
ReplaceAttribs = "[" & ReplaceAttribs & "]"
End Function
Sub EnrichirBaseFils(Attributes As Variant)
    Dim sql As String
    Dim SqlValues As String
    Dim sqlNull As Boolean
    Dim Rs As Recordset
    
    sqlNull = True
    Num = 0
    sql = "INSERT INTO xls_Ligne_Tableau_fils ( Job,"
    For i = LBound(Attributes) To UBound(Attributes)
        Debug.Print ReplaceAttribs(Attributes(i).TagString, sql)
        sql = sql & ReplaceAttribs(Attributes(i).TagString, sql) & ","
    Next i
    sql = Left(sql, Len(sql) - 1)
    sql = sql & ") Values ("
    SqlValues = ""
    
    SqlValues = SqlValues & NmJob & ","
    For i = LBound(Attributes) To UBound(Attributes)
    If UCase(Trim(Attributes(i).TextString)) = "FIL" Then Exit Sub
        If Trim("" & Attributes(i).TextString) = "" Then
            SqlValues = SqlValues & "NULL,"
        Else
            sqlNull = False
            SqlValues = SqlValues & "'" & Attributes(i).TextString & "',"
        End If
    Next i
    
    SqlValues = Left(SqlValues, Len(SqlValues) - 1)
      sql = sql & SqlValues & ");"
      If sqlNull = False Then Con.Exequte sql
End Sub
Sub EnrichirBaseConnecteur(Attributes As Variant, NameConnecteur)
    Dim sql As String
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
    
    sql = "INSERT INTO Xls_Connecteurs (Job,CONNECTEUR, [O/N], DESIGNATION,POS, N°,   CODE_APP, PRECO1, PRECO2 )"
    sql = sql & "Values ("
    SqlValues = ""
    
    SqlValues = SqlValues & NmJob & ",'" & NameConnecteur & "',false,"
    For i = LBound(Table) To UBound(Table)
        For i2 = LBound(Attributes) To UBound(Attributes)
            If Table(i) = PRECO(UCase("" & Attributes(i2).TagString)) Then
            
                If Trim("" & Attributes(i2).TextString) = "" Then
                    SqlValues = SqlValues & "NULL,"
                Else
                    SqlValues = SqlValues & "'" & Attributes(i2).TextString & "',"
                End If
                Exit For
            End If
       
        Next i2
    Next i
    SqlValues = Left(SqlValues, Len(SqlValues) - 1)
    sql = sql & SqlValues & ");"
    Debug.Print sql
    Con.Exequte sql
End Sub

Sub EnrichirBaseComposants(Attributes As Variant, NameConnecteur)
    Dim sql As String
    Dim SqlValues As String
    Dim Rs As Recordset
    Dim Table(3) As String
    
    Static Ip As Long
    Ip = Ip + 1
Table(0) = "DESIGNCOMP"
Table(1) = "NUMCOMP"
Table(2) = "REFCOMP"
Table(3) = "PATHCOMP"
    
    sql = "INSERT INTO Xls_Composants (Job,DESIGNCOMP, NUMCOMP, REFCOMP,Path )"
    sql = sql & "Values ("
    SqlValues = ""
    
    SqlValues = SqlValues & NmJob & ","
    For i = LBound(Table) To UBound(Table)
        For i2 = LBound(Attributes) To UBound(Attributes)
            If Table(i) = PRECO(UCase("" & Attributes(i2).TagString)) Then
            
                If Trim("" & Attributes(i2).TextString) = "" Then
                    SqlValues = SqlValues & "NULL,"
                Else
                    If Attributes(i2).TagString = "NUMCOMP" Then
                    
                        SqlValues = SqlValues & "" & CInt(Mid(Attributes(i2).TextString, 2, Len(Attributes(i2).TextString) - 1)) & ","
                    Else
                        SqlValues = SqlValues & "'" & Attributes(i2).TextString & "',"
                    End If
                End If
                Exit For
            End If
       
        Next i2
    Next i
    SqlValues = Left(SqlValues, Len(SqlValues) - 1)
    sql = sql & SqlValues & ");"
    Debug.Print sql
    Con.Exequte sql
End Sub
Sub EnrichirBaseNota(Attributes As Variant, NameConnecteur)
    Dim sql As String
    Dim SqlValues As String
    Dim Rs As Recordset
    Dim Table(0) As String
    
    Static Ip As Long
    Ip = Ip + 1

    Dim Trouve As Boolean
    
Table(0) = "NUMNOTA"
    sql = "INSERT INTO Xls_Nota (Job,NOTA, NUMNOTA )"
    sql = sql & "Values ("
    SqlValues = ""
    
    SqlValues = SqlValues & NmJob & ",'" & NameConnecteur & "',"
    For i = LBound(Table) To UBound(Table)
        For i2 = LBound(Attributes) To UBound(Attributes)
            If Table(i) = PRECO(UCase("" & Attributes(i2).TagString)) Then
            
                If Trim("" & Attributes(i2).TextString) = "" Then
                    SqlValues = SqlValues & "NULL,"
                Else
                
                    SqlValues = SqlValues & "'" & Attributes(i2).TextString & "',"
                End If
                Exit For
            End If
       
        Next i2
    Next i
    SqlValues = Left(SqlValues, Len(SqlValues) - 1)
    sql = sql & SqlValues & ");"
    Debug.Print sql
    Con.Exequte sql
End Sub
Sub EnrichirBaseConnecteurEpissure(Attributes As Variant, NameConnecteur)
    Dim sql As String
    Dim SqlValues As String
    Dim Rs As Recordset
    Dim Table(5) As String
   
  Table(0) = "DESIGNATION"

Table(1) = "N°"
Table(2) = "CODE_APP"
Table(3) = "POS"
Table(4) = "PRECO1"
Table(5) = "PRECO2"
    
    sql = "INSERT INTO Xls_Connecteurs (Job, CONNECTEUR, [O/N], DESIGNATION,  N°, CODE_APP,POS, PRECO1, PRECO2 )"
    sql = sql & "Values ("
    SqlValues = ""
        SqlValues = ""
        
    SqlValues = SqlValues & NmJob & ",'" & NameConnecteur & "',true,"

     For i = LBound(Table) To UBound(Table)
        For i2 = LBound(Attributes) To UBound(Attributes)
            If Table(i) = PRECO(UCase("" & Attributes(i2).TagString)) Then
            
                If Trim("" & Attributes(i2).TextString) = "" Then
                    SqlValues = SqlValues & "NULL,"
                Else
                    SqlValues = SqlValues & "'" & Attributes(i2).TextString & "',"
                End If
                Exit For
            End If
       
        Next i2
    Next i
    SqlValues = Left(SqlValues, Len(SqlValues) - 1)
    sql = sql & SqlValues & ");"
    Debug.Print sql
    Con.Exequte sql
End Sub
Sub MajImportAutocad(Id_IndiceProjet)
Dim sql As String
Dim SqlValues As String
Dim IConnecteur As Long
Dim boolReprise As Boolean
Dim Rs As Recordset
sql = "DELETE Connecteurs.*, Connecteurs.Id_IndiceProjet "
sql = sql & "FROM Connecteurs "
sql = sql & "WHERE Connecteurs.Id_IndiceProjet=" & Id_IndiceProjet & ";"

Con.Exequte sql

sql = "DELETE Ligne_Tableau_fils.*, Ligne_Tableau_fils.Id_IndiceProjet "
sql = sql & "FROM Ligne_Tableau_fils "
sql = sql & "WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & Id_IndiceProjet & ";"
Con.Exequte sql

sql = "SELECT " & Id_IndiceProjet & " AS Id_IndiceProjet, RqXls_Connecteurs.CONNECTEUR, RqXls_Connecteurs.[O/N], RqXls_Connecteurs.DESIGNATION, RqXls_Connecteurs.CODE_APP, RqXls_Connecteurs.N°, RqXls_Connecteurs.POS, RqXls_Connecteurs.PRECO1, RqXls_Connecteurs.PRECO2 "
sql = sql & "FROM RqXls_Connecteurs "
sql = sql & "WHERE RqXls_Connecteurs.Job=" & NmJob & ";"
Set Rs = Con.OpenRecordSet(sql)
IConnecteur = 0
While Rs.EOF = False

sql = "INSERT INTO Connecteurs ( Id_IndiceProjet, CONNECTEUR, [O/N], DESIGNATION, CODE_APP, N°, POS, PRECO1, PRECO2 ) "
sql = sql & "VALUES("
Reprise:
IConnecteur = IConnecteur + 1

SqlValues = Id_IndiceProjet & ","
If IConnecteur < Val("" & Rs![N°]) Then
    boolReprise = True
     SqlValues = SqlValues & "'NEANT',"
    For i = 2 To Rs.Fields.Count - 1
            If Rs.Fields(i).Name = "N°" Then
                SqlValues = SqlValues & "'" & IConnecteur & "',"
            Else
                SqlValues = SqlValues & "NULL,"
            End If
    Next i

Else
    boolReprise = False
For i = 1 To Rs.Fields.Count - 1
    If Trim("" & Rs.Fields(i)) = "" Then
        SqlValues = SqlValues & "NULL,"
    Else
        If Rs.Fields(i).Type = adBoolean Then
             SqlValues = SqlValues & Replace(Replace(Rs.Fields(i), "Faux", "false"), "Vrai", "true") & " ,"
        Else
        SqlValues = SqlValues & "'" & Trim("" & Rs.Fields(i)) & "',"
        End If
    End If
Next i
End If
    SqlValues = Left(SqlValues, Len(SqlValues) - 1)
    Con.Exequte sql & SqlValues & ");"
    If boolReprise = True Then GoTo Reprise
Rs.MoveNext
Wend

sql = "SELECT  Connecteurs.N° "
sql = sql & "FROM Connecteurs "
sql = sql & "WHERE Connecteurs.Id_IndiceProjet = " & Id_IndiceProjet & " "
sql = sql & "And Connecteurs.N° Is Not Null "
sql = sql & "ORDER BY Connecteurs.N° DESC;"
Set Rs = Con.OpenRecordSet(sql)
If Rs.EOF = False Then
NbCon = Rs!N°
sql = "SELECT  Connecteurs.Numéro "
sql = sql & "FROM Connecteurs "
sql = sql & "WHERE Connecteurs.Id_IndiceProjet = " & Id_IndiceProjet & " "
sql = sql & "And Connecteurs.N° Is Null "

Set Rs = Con.OpenRecordSet(sql)
While Rs.EOF = False
NbCon = NbCon + 1
sql = "UPDATE Connecteurs SET Connecteurs.N° = " & NbCon & " "
sql = sql & "WHERE Connecteurs.Numéro=" & Rs!Numéro & ";"
Con.Exequte sql
    Rs.MoveNext
Wend
End If



Set Rs = Con.CloseRecordSet(Rs)
sql = "INSERT INTO Ligne_Tableau_fils ( Id_IndiceProjet, LIAI, "
sql = sql & "DESIGNATION, FIL, SECT, TEINT, TEINT2, ISO, POS, "
sql = sql & "FA, VOI, POS2, FA2, VOI2, [LONG],APP,APP2 ) "
sql = sql & "SELECT " & Id_IndiceProjet & " AS Id_IndiceProjet, xls_Ligne_Tableau_fils.LIAI, "
sql = sql & "xls_Ligne_Tableau_fils.DESIGNATION, xls_Ligne_Tableau_fils.FIL, "
sql = sql & "xls_Ligne_Tableau_fils.SECT, xls_Ligne_Tableau_fils.TEINT, "
sql = sql & "xls_Ligne_Tableau_fils.TEINT2, xls_Ligne_Tableau_fils.ISO, "
sql = sql & "xls_Ligne_Tableau_fils.POS, xls_Ligne_Tableau_fils.FA, "
sql = sql & "xls_Ligne_Tableau_fils.VOI, xls_Ligne_Tableau_fils.POS2, "
sql = sql & "xls_Ligne_Tableau_fils.FA2, xls_Ligne_Tableau_fils.VOI2, "
sql = sql & "xls_Ligne_Tableau_fils.LONG, "
sql = sql & "(SELECT  Xls_Connecteurs.CODE_APP FROM Xls_Connecteurs WHERE Xls_Connecteurs.N°=[FA] and Xls_Connecteurs.Job=" & NmJob & ") as APP, "
sql = sql & "(SELECT  Xls_Connecteurs.CODE_APP FROM Xls_Connecteurs WHERE Xls_Connecteurs.N°=[FA2] and Xls_Connecteurs.Job=" & NmJob & ") as APP2 "

sql = sql & "FROM xls_Ligne_Tableau_fils "
sql = sql & "where xls_Ligne_Tableau_fils.Job=" & NmJob & ";"
Con.Exequte sql
End Sub
