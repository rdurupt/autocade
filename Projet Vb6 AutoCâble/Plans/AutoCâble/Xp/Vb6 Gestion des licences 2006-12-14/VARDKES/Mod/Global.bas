Attribute VB_Name = "Global"
Option Explicit
Public BdDateTable As String
Public DbNumPlan  As String

Global TableauPath As New Collection
Public TableauOnglet As New Collection
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

  
Open APP.Path & "\Autocable.ini" For Input As #FileNumber
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

Sub AjouterClassOnglet(ColecOnglet As Collection, Onglet As String)
Dim T_O As clsOnglet

 Set T_O = New clsOnglet
 T_O.Onglet = Onglet
    ColecOnglet.Add T_O, Onglet
     Set T_O = Nothing
    
End Sub
Public Function PathArchive(PathRacicine As String, Client As String, CleAc As String, Piece As String, Mytype As String, Fichier, IdPieces As Long, Indice_Pieces As String, Indice_Plan As String, Version As Long, Optional NoRegistre As Boolean) As String
Dim Fso As New FileSystemObject
Dim Sql As String
Dim RS As Recordset
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
    Dim RS As Recordset
        Set RS = Con.OpenRecordSet("SELECT T_Path.* FROM T_Path;")
    While RS.EOF = False
        MyPath.Add RS.Fields("PathVar").Value, RS.Fields("NameVar").Value
        RS.MoveNext
    Wend
    Set RS = Con.CloseRecordSet(RS)
    Set funPath = MyPath
End Function

Sub Main()

If Trim("" & Command) <> "[N ais pas peur fils papa est la]" Then
    MsgBox "Ce module ne peut être exécuté qu'à partir d'une licence Autocâble.", vbCritical, "VARDKES Automation & Co"
    End
End If
LoadDb
ConeverEtudeCsv.Show vbModal

Con.CloseConnection

End Sub
