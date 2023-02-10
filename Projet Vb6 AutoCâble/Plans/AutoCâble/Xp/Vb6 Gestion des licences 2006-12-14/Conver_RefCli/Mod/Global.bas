Attribute VB_Name = "Global"
Option Explicit
Public BdDateTable As String
Public DbNumPlan  As String
Public NmJob As Long
Global TableauPath As New Collection

Public ADO_TYPEBASE  As String
Public ADO_BASE  As String
Public ADO_SERVER As String
Public ADO_Fichier As String
Public ADO_User As String
Public ADO_PassWord As String
Public AutocableDRIVE  As String
Public DonneesEntreprise  As String
Public DonneesProduction  As String
Public IsCilent As Boolean
Public IsServeur As Boolean
Public Con As New Ado
Dim bool_MiseEnPage As Boolean
Sub DeletSheet(MySheet As Excel.Worksheet)
    On Error Resume Next
    MySheet.Delete
Err.Clear
End Sub
Function LaodJob() As Long
Dim Sql As String
Dim Rs As Recordset
If NmJob = 0 Then

Sql = "SELECT [NumErreur]+1 AS Job FROM T_NumErreur WHERE T_NumErreur.LibErreur='Job';"
Set Rs = Con.OpenRecordSet(Sql)
LaodJob = Rs!Job
Sql = "UPDATE T_NumErreur SET T_NumErreur.NumErreur = [NumErreur]+1 WHERE T_NumErreur.LibErreur='Job';"
Con.Execute Sql
Set Rs = Con.CloseRecordSet(Rs)

Else
    LaodJob = NmJob
End If
End Function
Function MyReplace(strVal As String) As String
strVal = Trim(strVal)
MyReplace = strVal
MyReplace = Replace(MyReplace, "'", "''")
MyReplace = Replace(MyReplace, Chr(34), Chr(34) & Chr(34))
MyReplace = Trim("" & MyReplace)
End Function
Function funOpenDatabase()
Con.TYPEBASE = ADO_TYPEBASE
Con.BASE = ADO_BASE
Con.SERVER = ADO_SERVER
Con.Fichier = ADO_Fichier
Con.User = ADO_User
Con.PassWord = ADO_PassWord
Con.OpenConnetion
End Function
Sub LoadDb()

BdDateTable = CherCheInFihier("BdDateTable")
DbNumPlan = CherCheInFihier("Bdnumero")
If UCase(CherCheInFihier("IsCilent")) = "TRUE" Then IsCilent = True

If UCase(CherCheInFihier("IsServeur")) = "TRUE" Then IsServeur = True
ADO_TYPEBASE = CherCheInFihier("ADO_TYPEBASE")
ADO_BASE = CherCheInFihier("ADO_BASE")
ADO_SERVER = CherCheInFihier("ADO_SERVER")
ADO_Fichier = CherCheInFihier("ADO_Fichier")
ADO_User = CherCheInFihier("ADO_User=")
ADO_PassWord = CherCheInFihier("ADO_PassWord")
AutocableDRIVE = CherCheInFihier("AutocableDRIVE")
DonneesEntreprise = CherCheInFihier("DonneesEntreprise")
DonneesProduction = CherCheInFihier("DonneesProduction")
funOpenDatabase

If IsServeur = IsCilent Then IsServeur = False: IsCilent = False
End Sub
Function CherCheInFihier(Cherher As String) As String
Dim FileNumber As Long
Dim MyString As String
Dim Spliligne
FileNumber = FreeFile

  
Open App.Path & "\Autocable.ini" For Input As #FileNumber
Do While Not EOF(FileNumber)    ' Effectue la boucle jusqu'à la fin du fichier.
    Input #FileNumber, MyString ' Lit les données dans deux variables.
    If InStr(1, MyString, Cherher) <> 0 Then
    Spliligne = Split(MyString & "====", "=")
       CherCheInFihier = Trim(Spliligne(1))
'       CherCheInFihier = Right(MyString, Len(MyString) - (Len(Cherher) + InStr(1, MyString, Cherher)))
       Exit Do
    End If
Loop
Close #FileNumber    ' Ferme le fichier
CherCheInFihier = Replace(CherCheInFihier, "§remplaceDate§", Format(Date, "yyyy"))
CherCheInFihier = Trim(CherCheInFihier)
End Function

Public Function PathArchive(PathRacicine As String, Client As String, CleAc As String, Piece As String, Mytype As String, Fichier, IdPieces As Long, Indice_Pieces As String, Indice_Plan As String, Version As Long, Optional NoRegistre As Boolean) As String
Dim Fso As New FileSystemObject
Dim Sql As String
Dim Rs As Recordset
Dim MyPath
Dim aa
Dim IndexP As Long
Indice_Pieces = Trim("" & Indice_Pieces)
Indice_Plan = Trim("" & Indice_Plan)
Piece = Replace(Piece, "/", "_", 1)
Piece = Replace(Piece, ":", "", 1)
Piece = Replace(Piece, ".", "", 1)
Piece = Piece & "_" & Indice_Pieces
If UCase(Mytype) = UCase("SyntG") Or UCase(Mytype) = UCase("pdf") Or UCase(Mytype) = UCase("Synt") Or Mytype = "LIEC" Or Mytype = "DAC" Or Mytype = "DNC" Or Mytype = "FAB" Then
Else
Fichier = Fichier & "_" & Indice_Plan
End If
Fichier = Replace(Fichier, "/", "_", 1)
Fichier = Replace(Fichier, ":", "", 1)
Fichier = Replace(Fichier, ".", "", 1)



PathArchive = TableauPath.Item(Mytype)
PathArchive = Replace(UCase(PathArchive), UCase("[VariableClient]"), Client)
PathArchive = Replace(UCase(PathArchive), UCase("[VariableAff]"), CleAc)
    PathArchive = Replace(UCase(PathArchive), UCase("[VaribleDoc]"), Fichier)

    


If Version > 1 Then
    
    PathArchive = Replace(UCase(PathArchive), UCase("[VARIABLEPI]"), Piece & "_MOD")
Else
    PathArchive = Replace(UCase(PathArchive), UCase("[VARIABLEPI]"), Piece)
End If

PathRacicine = DefinirChemienComplet(TableauPath.Item("PathServer"), PathRacicine)
MyPath = Split(PathArchive, "\")
aa = ""
For IndexP = 0 To UBound(MyPath) - 1
aa = aa & MyPath(IndexP) & "\"
Debug.Print Left(PathRacicine & "\" & aa, Len(PathRacicine & "\" & aa) - 1)
If Fso.FolderExists(Left(PathRacicine & "\" & aa, Len(PathRacicine & "\" & aa) - 1)) = False Then
    Fso.CreateFolder Left(PathRacicine & "\" & aa, Len(PathRacicine & "\" & aa) - 1)
End If
Next



    


PathArchive = PathRacicine & "\" & PathArchive
Debug.Print PathArchive
End Function
Function DefinirChemienComplet(Serveur As String, Path As String) As String
If Right(Trim("" & Serveur), 1) <> "\" Then Serveur = Serveur & "\"
If Trim("" & Path) = "" Then
    DefinirChemienComplet = Serveur
Else
    If Left(Path, 1) = "\" And Left(Path, 2) <> "\\" Then Path = Right(Path, Len(Path) - 1)
DefinirChemienComplet = Path
End If
If Mid(DefinirChemienComplet, 2, 1) = ":" Then Exit Function
If Left(Path, 1) <> "\" Then
    If Right(Serveur, 1) <> "\" Then
        DefinirChemienComplet = Serveur & "\" & DefinirChemienComplet
    Else
         DefinirChemienComplet = Serveur & DefinirChemienComplet
    End If
End If
If Right(Trim(DefinirChemienComplet), 2) = "\\" Then DefinirChemienComplet = Mid(DefinirChemienComplet, 1, Len(DefinirChemienComplet) - 1)
If Left(DefinirChemienComplet, 1) = "\" And Left(DefinirChemienComplet, 2) <> "\\" Then DefinirChemienComplet = "\" & DefinirChemienComplet
Debug.Print DefinirChemienComplet
End Function
Function funPath()
    Dim MyPath As New Collection
    Dim Rs As Recordset
        Set Rs = Con.OpenRecordSet("SELECT T_Path.* FROM T_Path;")
    While Rs.EOF = False
        MyPath.Add Rs.Fields("PathVar").Value, Rs.Fields("NameVar").Value
        Rs.MoveNext
    Wend
    Set Rs = Con.CloseRecordSet(Rs)
    Set funPath = MyPath
End Function

Sub Main()

If Trim("" & Command) <> "[N ais pas peur fils papa est la]" Then

    MsgBox "Ce module ne peut être exécuté qu'à partir d'une licence Autocâble.", vbCritical, "Conversion CLI"

    End
End If
LoadDb

ConeverEtudeCsv.Show vbModal

Con.CloseConnection

End Sub
Sub MajEcart(IdIndiceProjet As Long, IdFils As Long, MyExcel As Excel.Application)
Dim PathPl As String
Dim PathPl2 As String
Dim MyFormatDate As String
 Set TableauPath = funPath
Dim L As Long
Dim C As Long
Dim boolSave As Boolean
Dim Sql As String
Dim RsSuprimer As Recordset
Dim RsAjouter As Recordset
Dim RsModifier As Recordset
Dim PathArchiveAutocad As String
Dim MyWorkbook As Workbook
Dim I As Long
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
Sub Racourci(RaccourciName As String, RaccourciCible As String, Extension As String)
Dim objshell
Dim objraccourci
Dim Fso As New FileSystemObject
If Fso.FileExists(RaccourciName & ".Lnk") = True Then
     Fso.DeleteFile RaccourciName & ".Lnk"
End If
Set objshell = CreateObject("wscript.shell")
Set objraccourci = objshell.createshortcut(RaccourciName & ".Lnk")
objraccourci.targetpath = RaccourciCible & "." & Extension
objraccourci.Save
Set Fso = Nothing
Set objraccourci = Nothing
End Sub
Sub RecherModifier(RsModifier As Recordset)
Dim I As Long
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
Public Function IsertSheet(MyWorkbook As Excel.Workbook, Name As String, Optional Fin As Boolean) As Excel.Worksheet
On Error Resume Next
Name = Trim(Name)
If Trim(Name) = "Appro Connectique" Then
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
 
 Name = Trim(Name) & Space(31)
 IsertSheet.Select
 IsertSheet.Name = Trim(Left(Name, 31))
 On Error GoTo 0
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


Function MajEcartExcel(IdIndiceProjet As Long, MyWorkbook As Excel.Workbook, L As Long, C As Long, RsSuprimer As Recordset, RsAjouter As Recordset, RsModifier As Recordset, SheetName As String, Optional Txt = "") As Boolean

Dim PortraitPaysage As Long
Dim modifire As Boolean
Dim a
Dim aa
Dim Sql As String
Dim TruveSheet As Boolean
Dim boolTxt As Boolean
Dim MySheet As Excel.Worksheet
Dim MyRange As Range
Dim T_Txt
Dim I2 As Long
Dim I As Long
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

Sub MiseEnPage(MyWorksheet As Worksheet, MyRange As Range, MyLeftHeader As String, _
            MyCenterHeader As String, MyRightHeader As String, MyLeftFooter As String, _
            MyCenterFooter As String, MyRightFooter As String, _
            MyZoom, CellVolet As String, RepeatCol As Boolean, MyxlLandscape As Long, _
            Optional AutoFilterOk As Boolean, Optional NotCouleur As Boolean, Optional MergeOk As Boolean, _
            Optional BottomMargin As Double = 2.5, Optional AutoFit As Boolean = True, Optional ZoneImpression As Boolean = True)
'            MyWorksheet.Application.Visible = True
'
Dim aa
Dim C As Long
On Error Resume Next
            MyWorksheet.Select
          If Trim(CellVolet) <> "" Then
  MyWorksheet.Range(CellVolet).Select
  End If
  If AutoFit = True Then
        MyWorksheet.Cells.ColumnWidth = 255
        MyWorksheet.Cells.RowHeight = 255
        MyWorksheet.Cells.EntireRow.AutoFit
        MyWorksheet.Cells.EntireColumn.AutoFit
        
    End If
 
        MyWorksheet.Range(MyRange(1, 1).Address & ":" & MyRange(1, MyRange.Columns.Count).Address).HorizontalAlignment = xlCenter
'        .VerticalAlignment = xlContext
        MyWorksheet.Range(MyRange(1, 1).Address & ":" & MyRange(1, MyRange.Columns.Count).Address).WrapText = True
        MyWorksheet.Range(MyRange(1, 1).Address & ":" & MyRange(1, MyRange.Columns.Count).Address).Orientation = 0
        MyWorksheet.Range(MyRange(1, 1).Address & ":" & MyRange(1, MyRange.Columns.Count).Address).AddIndent = False
        MyWorksheet.Range(MyRange(1, 1).Address & ":" & MyRange(1, MyRange.Columns.Count).Address).IndentLevel = 0
        MyWorksheet.Range(MyRange(1, 1).Address & ":" & MyRange(1, MyRange.Columns.Count).Address).ShrinkToFit = False
        MyWorksheet.Range(MyRange(1, 1).Address & ":" & MyRange(1, MyRange.Columns.Count).Address).ReadingOrder = xlContext
        MyWorksheet.Range(MyRange(1, 1).Address & ":" & MyRange(1, MyRange.Columns.Count).Address).MergeCells = MergeOk
   
    If NotCouleur = False Then _
    MyWorksheet.Range(MyRange(1, 1).Address & ":" & MyRange(1, MyRange.Columns.Count).Address).Interior.ColorIndex = 15
    
    If AutoFit = True Then
        MyWorksheet.Cells.ColumnWidth = 255
        MyWorksheet.Cells.RowHeight = 255
        MyWorksheet.Cells.EntireRow.AutoFit
        MyWorksheet.Cells.EntireColumn.AutoFit
        
    End If
'  MyWorksheet.Application.Visible = True
 If Trim(CellVolet) <> "" Then
  MyWorksheet.Application.ActiveWindow.FreezePanes = True
  End If
  If bool_MiseEnPage = True Then
  If ZoneImpression = True Then
        MyWorksheet.PageSetup.PrintArea = "A1:" & MyRange(MyRange.Rows.Count, MyRange.Columns.Count).Address
    End If
     
    
    MyWorksheet.PageSetup.LeftHeader = Replace(MyLeftHeader, vbCrLf, Chr(10))
    DoEvents
     MyWorksheet.PageSetup.CenterHeader = Replace(MyCenterHeader, vbCrLf, Chr(10))
   DoEvents
   MyWorksheet.PageSetup.RightHeader = Replace(MyRightHeader, vbCrLf, Chr(10))
    DoEvents
   
    
    MyWorksheet.PageSetup.TopMargin = MyWorksheet.Application.InchesToPoints(2)
    MyWorksheet.PageSetup.LeftMargin = MyWorksheet.Application.InchesToPoints(0.196850393700787)
    MyWorksheet.PageSetup.RightMargin = MyWorksheet.Application.InchesToPoints(0.196850393700787)
    MyWorksheet.PageSetup.TopMargin = MyWorksheet.Application.InchesToPoints(1.37795275590551)
    MyWorksheet.PageSetup.BottomMargin = MyWorksheet.Application.InchesToPoints(BottomMargin / 2.54)  '0.984251968503937)
    MyWorksheet.PageSetup.HeaderMargin = MyWorksheet.Application.InchesToPoints(0.196850393700787)
     aa = 0.5 / 2.54
     Debug.Print aa
    MyWorksheet.PageSetup.FooterMargin = MyWorksheet.Application.InchesToPoints(0.196850393700787)  '0.984251968503937)
     
    MyWorksheet.PageSetup.LeftFooter = Replace(MyLeftFooter, Chr(13), "")
    MyWorksheet.PageSetup.CenterFooter = Replace(MyCenterFooter, Chr(13), "")
    MyWorksheet.PageSetup.RightFooter = Replace(MyRightFooter, Chr(13), "")
    MyWorksheet.PageSetup.Orientation = MyxlLandscape
    MyWorksheet.PageSetup.Draft = False
    MyWorksheet.PageSetup.PaperSize = xlPaperA4
    MyWorksheet.PageSetup.FirstPageNumber = xlAutomatic
    MyWorksheet.PageSetup.Order = xlDownThenOver
    MyWorksheet.PageSetup.BlackAndWhite = False
    MyWorksheet.PageSetup.Zoom = MyZoom
    MyWorksheet.PageSetup.FitToPagesWide = 1
    MyWorksheet.PageSetup.FitToPagesTall = 1
    MyWorksheet.PageSetup.PrintErrors = xlPrintErrorsDisplayed
    MyWorksheet.PageSetup.CenterHorizontally = True
    MyWorksheet.PageSetup.PrintGridlines = False
    
      
    MyWorksheet.PageSetup.PrintTitleRows = MyWorksheet.Range(MyRange(1, 1).Address & ":" & MyRange(1, MyRange.Columns.Count).Address).Address
    
     If RepeatCol = True Then _
     MyWorksheet.PageSetup.PrintTitleColumns = MyWorksheet.Range(MyRange(1, 1).Address & ":" & MyRange(1, 1).Address).Address
     
    End If
           
           
   
End Sub
