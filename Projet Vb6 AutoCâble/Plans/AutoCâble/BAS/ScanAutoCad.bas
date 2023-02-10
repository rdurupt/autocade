Attribute VB_Name = "ScanAutoCad"

Public Sub ScanDessin(Fichier As String, Projet As String, Indce As String, Description As String, LI As String, CleAc)
    Dim Fso As New FileSystemObject
    Dim NewBlock  As AcadBlock
    Dim NewBlock2  As AcadBlockReference
    Dim Entity As AcadEntity
    Dim BlocRef As AcadBlockReference
    Dim Attributes As Variant
   Dim Sql As String
Dim Rs As Recordset
Dim IdProjet As Long
Dim IdIndice As Long
DoEvents
Con.OpenConnetion db
Sql = "SELECT T_Projet.id FROM T_Projet WHERE T_Projet.Projet='" & Projet & "';"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = False Then
  IdProjet = Rs!Id
Else
Sql = "INSERT INTO T_Projet ( Projet,CleAc )"
Sql = Sql & "Values('" & MyReplace(Projet) & "'," & CleAc & ");"
Con.Exequte Sql
Rs.Requery
 IdProjet = Rs!Id
End If

Sql = "SELECT T_indiceProjet.id FROM T_indiceProjet WHERE T_indiceProjet.LI='" & MyReplace(LI) & "' and T_indiceProjet.IdProjet=" & IdProjet & ";"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = False Then
  IdIndice = Rs!Id
Else
Sql = "INSERT INTO T_indiceProjet ( IdProjet, Indice, Description ,LI)"
Sql = Sql & "values( " & IdProjet & " , '" & Indce & "', '" & MyReplace(Description) & "','" & MyReplace(LI) & "' );"
Con.Exequte Sql
Rs.Requery
IdIndice = Rs!Id
End If
    Con.Exequte "DELETE xls_Ligne_Tableau_fils.* FROM xls_Ligne_Tableau_fils;"
    
    Set AutoApp = ThisDrawing.Application
    OpenFichier Fichier
    Menu.ProgressBar1Caption = "Lecture Tableau des Fils:"
    Menu.ProgressBar1.Value = 0
    Menu.ProgressBar1.Max = AutoApp.ActiveDocument.ModelSpace.Count
    
    For i = 0 To AutoApp.ActiveDocument.ModelSpace.Count - 1
    Menu.ProgressBar1.Value = Menu.ProgressBar1.Value + 1
    DoEvents
        Set Entity = AutoApp.ActiveDocument.ModelSpace.Item(i)
       
        If Entity.ObjectName = "AcDbBlockReference" Then
            Set BlocRef = Entity
            a = BlocRef.Name
            If BlocRef.HasAttributes Then
                Attributes = BlocRef.GetAttributes
                If (UBound(Attributes) = 13) Or (UBound(Attributes) = 12) Then
                    If IsTableauFils(Attributes) = True Then
                        EnrichirBaseFils Attributes
                    End If
                End If
            End If
        End If
    Next i

    Con.Exequte "DELETE Xls_Connecteurs.* FROM Xls_Connecteurs;"

      Menu.ProgressBar1Caption = "Lecture des Connecteurs:"
    Menu.ProgressBar1.Value = 0
    Menu.ProgressBar1.Max = AutoApp.ActiveDocument.ModelSpace.Count
    
    For i = 0 To AutoApp.ActiveDocument.ModelSpace.Count - 1
    Menu.ProgressBar1.Value = Menu.ProgressBar1.Value + 1
    DoEvents
        Set Entity = AutoApp.ActiveDocument.ModelSpace.Item(i)
        If Entity.ObjectName = "AcDbBlockReference" Then
            Set BlocRef = Entity
            If BlocRef.HasAttributes Then
           a = BlocRef.Name
         
                Attributes = BlocRef.GetAttributes
'                If UBound(Attributes) = 13 Then
                    If IsConnecteurs(Attributes) = True Then
                        EnrichirBaseConnecteur Attributes, BlocRef.Name
                    Else
                        If IsEpissures(Attributes) = True Then
                            EnrichirBaseConnecteurEpissure Attributes, BlocRef.Name
                        End If
                    End If
'                End If
            End If
        End If
    Next i
    MajImportAutocad IdIndice
    Con.CloseConnection
    Menu.ProgressBar1Caption = "Traitement terminé:"
    Menu.ProgressBar1.Value = 0
   CloseDocument
    End Sub
    


Function PRECO(Var As String) As String
PRECO = Var
PRECO = Replace(PRECO, "CODE.APP", "CODE_APP")
PRECO = Replace(PRECO, "FILA", "FIL")
PRECO = Replace(PRECO, "FILB", "FIL")
PRECO = Replace(PRECO, "FIL1", "FIL")
PRECO = Replace(PRECO, "FILG1", "FILG")

If InStr(1, PRECO, "PRECO") <> 0 Then
    PRECO = "PRECO" & Right(PRECO, 1)
    
End If
End Function
Sub EnrichirBaseFils(Attributes As Variant)
    Dim Sql As String
    Dim SqlValues As String
    Dim sqlNull As Boolean
    Dim Rs As Recordset
    sqlNull = True
    Sql = "INSERT INTO xls_Ligne_Tableau_fils ( LIAI, DESIGNATION, FIL, SECT, TEINT, TEINT2, ISO, POS, FA, VOI, POS2, FA2, VOI2, [LONG] )"
    Sql = Sql & "Values ("
    SqlValues = ""
    For i = LBound(Attributes) To UBound(Attributes)
        If Trim("" & Attributes(i).TextString) = "" Then
            SqlValues = SqlValues & "NULL,"
        Else
            sqlNull = False
            SqlValues = SqlValues & "'" & Attributes(i).TextString & "',"
        End If
    Next i
    If i = 13 Then SqlValues = SqlValues & "NULL,"
    SqlValues = Left(SqlValues, Len(SqlValues) - 1)
      Sql = Sql & SqlValues & ");"
      If sqlNull = False Then Con.Exequte Sql
End Sub
Sub EnrichirBaseConnecteur(Attributes As Variant, NameConnecteur)
    Dim Sql As String
    Dim SqlValues As String
    Dim Rs As Recordset
    Dim Table(5) As String
    
    Static Ip As Long
'    Ip = 0
    Ip = Ip + 1
'    If Ip = 19 Then
'        MsgBox ""
'    End If
Table(0) = "DESIGNATION"
Table(1) = "POS"
Table(2) = "N°"
Table(3) = "CODE_APP"
Table(4) = "PRECO1"
Table(5) = "PRECO2"
'For i = LBound(Attributes) To UBound(Attributes)
'    Debug.Print Attributes(i).TagString & " : "; Attributes(i).TextString
'Next i
    
    Sql = "INSERT INTO Xls_Connecteurs ( CONNECTEUR, [EPISSURE O/N], DESIGNATION, CODE_APP, N°, POS, PRECO1, PRECO2 )"
    Sql = Sql & "Values ("
    SqlValues = ""
    SqlValues = SqlValues & "'" & NameConnecteur & "','N',"
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
    Sql = Sql & SqlValues & ");"
    Debug.Print Sql
    Con.Exequte Sql
End Sub
Sub EnrichirBaseConnecteurEpissure(Attributes As Variant, NameConnecteur)
    Dim Sql As String
    Dim SqlValues As String
    Dim Rs As Recordset
    Dim Table(5) As String
   
  Table(0) = "DESIGNATION"
Table(1) = "POS"
Table(2) = "N°"
Table(3) = "CODE_APP"
Table(4) = "PRECO1"
Table(5) = "PRECO2"
'Table(7) = "LIAI1"



    
    Sql = "INSERT INTO Xls_Connecteurs ( CONNECTEUR, [EPISSURE O/N], DESIGNATION, CODE_APP, N°, POS, PRECO1, PRECO2 )"
    Sql = Sql & "Values ("
    SqlValues = ""
        SqlValues = ""
    SqlValues = SqlValues & "'" & NameConnecteur & "','O',"

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
    Sql = Sql & SqlValues & ");"
    Debug.Print Sql
    Con.Exequte Sql
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

Con.Exequte Sql

Sql = "DELETE Ligne_Tableau_fils.*, Ligne_Tableau_fils.Id_IndiceProjet "
Sql = Sql & "FROM Ligne_Tableau_fils "
Sql = Sql & "WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & Id_IndiceProjet & ";"
Con.Exequte Sql

Sql = "SELECT " & Id_IndiceProjet & " AS Id_IndiceProjet, RqXls_Connecteurs.CONNECTEUR, RqXls_Connecteurs.[EPISSURE O/N], RqXls_Connecteurs.DESIGNATION, RqXls_Connecteurs.CODE_APP, RqXls_Connecteurs.N°, RqXls_Connecteurs.POS, RqXls_Connecteurs.PRECO1, RqXls_Connecteurs.PRECO2 "
Sql = Sql & "FROM RqXls_Connecteurs;"
Set Rs = Con.OpenRecordSet(Sql)
IConnecteur = 0
While Rs.EOF = False

Sql = "INSERT INTO Connecteurs ( Id_IndiceProjet, CONNECTEUR, [EPISSURE O/N], DESIGNATION, CODE_APP, N°, POS, PRECO1, PRECO2 ) "
Sql = Sql & "VALUES("
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
        SqlValues = SqlValues & "'" & Trim("" & Rs.Fields(i)) & "',"
    End If
Next i
End If
    SqlValues = Left(SqlValues, Len(SqlValues) - 1)
    Con.Exequte Sql & SqlValues & ");"
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
'Sql = Sql & "ORDER BY Connecteurs.N° DESC;"
Set Rs = Con.OpenRecordSet(Sql)
While Rs.EOF = False
NbCon = NbCon + 1
Sql = "UPDATE Connecteurs SET Connecteurs.N° = " & NbCon & " "
Sql = Sql & "WHERE Connecteurs.Numéro=" & Rs!Numéro & ";"
Con.Exequte Sql
    Rs.MoveNext
Wend
End If



Set Rs = Con.CloseRecordSet(Rs)
Sql = "INSERT INTO Ligne_Tableau_fils ( Id_IndiceProjet, LIAI, "
Sql = Sql & "DESIGNATION, FIL, SECT, TEINT, TEINT2, ISO, POS, "
Sql = Sql & "FA, VOI, POS2, FA2, VOI2, [LONG] ) "
Sql = Sql & "SELECT " & Id_IndiceProjet & " AS Id_IndiceProjet, xls_Ligne_Tableau_fils.LIAI, "
Sql = Sql & "xls_Ligne_Tableau_fils.DESIGNATION, xls_Ligne_Tableau_fils.FIL, "
Sql = Sql & "xls_Ligne_Tableau_fils.SECT, xls_Ligne_Tableau_fils.TEINT, "
Sql = Sql & "xls_Ligne_Tableau_fils.TEINT2, xls_Ligne_Tableau_fils.ISO, "
Sql = Sql & "xls_Ligne_Tableau_fils.POS, xls_Ligne_Tableau_fils.FA, "
Sql = Sql & "xls_Ligne_Tableau_fils.VOI, xls_Ligne_Tableau_fils.POS2, "
Sql = Sql & "xls_Ligne_Tableau_fils.FA2, xls_Ligne_Tableau_fils.VOI2, "
Sql = Sql & "xls_Ligne_Tableau_fils.LONG "
Sql = Sql & "FROM xls_Ligne_Tableau_fils;"
Con.Exequte Sql
End Sub
Sub TestScan()
'ScanDessin "C:\RD\TestPlan\D4F ind D.dwg"
' ScanDessin "C:\RD\TestPlan\F4R_RS_ind_A.dwg"
'ScanDessin "C:\RD\TestPlan\BRV.dwg"
    ScanDessin "C:\RD\TestPlan\rd.dwg"
MsgBox "Traitement terminé"
End Sub
