VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GenerateurDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Menu As String
Public PathDoc As String
Public DapresDocs
Private DapresDoc() As String
Private indexdoc As Long
Public id_Menu As Long
Public DocExecuter As Boolean
Public SurDoc As Boolean
Public SaveAs As String
Public IndexTableauDoc As Long
Public IsPDF As Boolean
Public MyClonne As Collection
Public FiltreMyClonne As Collection

Public MyColonneAs As Collection
Public Onglet As String
Private LstOnglet As String
Private ChamName
Public LstOngletName As String
Private MyExcel As EXCEL.Application
Private Type PlageOnget
            Debut As Long
            Fin As Long
            Err As Boolean
        End Type
Private MyPlageOnglets As PlageOnget
Private RsChampMasquer As Recordset
Public OfseteMiseEnPage As Long
Private Sub RenameTitre(MySeet As Worksheet)
Dim MyRange As Range
Dim C As Long
Set MyRange = MySeet.Range("A1").CurrentRegion
For C = 1 To MyClonne.Count
    MyRange(1, MyClonne(C)) = MyColonneAs(C)
Next
End Sub
Private Sub RensiegneColoneFilter(MySeet As Worksheet)
Dim MyRange As Range
Dim C As Long
Set FiltreMyClonne = Nothing
Set FiltreMyClonne = New Collection
Set MyRange = MySeet.Range("A1").CurrentRegion
For C = 1 To MyClonne.Count
    FiltreMyClonne.Add C, MyRange(1, MyClonne(C))
Next
End Sub
Private Function RetounSheet(MyWorkbook As Workbook, sheetSource As Worksheet, Id_Onglet As Long, Optional Prefix As String)
Dim MySeet As Worksheet
Dim Sql As String
Dim MyRange As Range
Dim Trie
Dim MyTrie
Dim MyTrieS
Dim I, I2, Col As Long
Dim NbTrie As Integer
Dim TrieExcel(3, 1) As String
'If Prefix = "C_" Then MsgBox "C_"
If Left(Trim(UCase(sheetSource.Name)), Len(Trim(UCase(Prefix))) - 1) = Left(Trim(UCase(Prefix)), Len(Trim(UCase(Prefix))) - 1) Then
    Prefix = ""
End If
Set MySeet = IsertSheet(MyWorkbook, Prefix & sheetSource.Name)
MySeet.Delete
Set MySeet = IsertSheet(MyWorkbook, Prefix & sheetSource.Name)
 Sql = "SELECT DISTINCT T_Etats_Onglet.Id AS Id_Select, T_Etats_Select_Champs.ChamsName, T_Etats_Select_Champs.Visible "
    Sql = Sql & "FROM T_Etats_Onglet INNER JOIN T_Etats_Select_Champs ON T_Etats_Onglet.Id = T_Etats_Select_Champs.Id_Onglet "
    Sql = Sql & "Where T_Etats_Onglet.Id = " & Id_Onglet & " "
    Sql = Sql & "And T_Etats_Select_Champs.Visible = false; "
    Set RsChampMasquer = Con.OpenRecordSet(Sql)
Set MyRange = sheetSource.Range("A1").CurrentRegion

MyRange.Copy
MyWorkbook.Activate
MySeet.Select
MySeet.Range("A1").Select
MySeet.Paste
RenseigneColonn MySeet, Id_Onglet, 1
 MyTrie = Split(Trie & "§", "§")
 If UBound(MyTrie) > 1 Then
    For I = 1 To UBound(MyTrie)
        MyTrieS = Split(MyTrie(I) & "=", "=")
    If Val(MyTrieS(1)) <> 0 Then
    TrieExcel(I, 0) = MyTrieS(0): TrieExcel(I, 1) = MyTrieS(1)
    If I = 3 Then Exit For
    End If
   
 Next
 Set MyRange = MySeet.Range("A1").CurrentRegion
NbTrie = 0
 For I = 1 To 3
    For Col = 1 To MyRange.Columns.Count
       If MyRange(1, Col) = TrieExcel(I, 0) Then
       NbTrie = NbTrie + 1
           TrieExcel(I, 0) = MyRange(2, Col).Address
            TrieExcel(I, 0) = Replace(TrieExcel(I, 0), "$", "")
           Exit For
       End If
    Next
  Next
 End If
 Trier MySeet, NbTrie, Replace("A2:" & MyRange(MyRange.Rows.Count, MyRange.Columns.Count).Address, "$", ""), TrieExcel(1, 0), Val(TrieExcel(1, 1)), TrieExcel(2, 0), Val(TrieExcel(2, 1)), TrieExcel(3, 0), Val(TrieExcel(3, 1))
 Set MyRange = MySeet.Range("A1").CurrentRegion
 For I = 1 To MyRange.Columns.Count
    If UCase(MyRange(1, I)) = "OPTION" Or UCase(MyRange(1, I)) = UCase("Criteres") Then
        For I2 = 2 To MyRange.Rows.Count
            MyRange(I2, I) = ";" & MyRange(I2, I) & ";"
            
        Next
        Exit For
    End If
 Next
 
 Set CollecFiltreName = Nothing
 Set CollecFiltreName = New Collection
 Set RetounSheet = MySeet
End Function

Public Sub RenseigneColonn(MySheet As Worksheet, id_Menu As Long, start As Long)
Dim L As Long
Dim C As Long
Dim Rs As Recordset
Dim MyRange As Range
Dim Sql As String
Set MyClonne = New Collection
Set MyColonneAs = New Collection
Set MyRange = MySheet.Range("A" & CStr(start)).CurrentRegion
For C = 1 To MyRange.Columns.Count
    Sql = "SELECT T_Etats_Onglet.Id, T_Etats_Select_Champs.ChamsName,T_Etats_Select_Champs.ChampAs, "
    Sql = Sql & "T_Etats_Select_Champs.CreatOnglet "
    Sql = Sql & "FROM T_Etats_Onglet INNER JOIN T_Etats_Select_Champs ON T_Etats_Onglet.Id =  "
    Sql = Sql & "T_Etats_Select_Champs.Id_Onglet "
    Sql = Sql & "WHERE (T_Etats_Select_Champs.ChamsName='" & Replace("" & MyRange(1, C), "'", "''") & "'  "
    Sql = Sql & "OR T_Etats_Select_Champs.ChampAs='" & Replace("" & MyRange(1, C), "'", "''") & "'  ) "
    Sql = Sql & "AND T_Etats_Onglet.Id=" & id_Menu & ";"
    Set Rs = Con.OpenRecordSet(Sql)
    If Rs.EOF = False Then
        If UCase(Trim("" & Rs!ChamsName)) = UCase(MyRange(1, C)) Then
            MyClonne.Add C, MyRange(1, C)
            MyColonneAs.Add Trim("" & Rs!ChampAs), MyRange(1, C)
        Else
            MyClonne.Add C, Trim("" & Rs!ChamsName)
            MyColonneAs.Add Trim("" & Rs!ChampAs), Trim("" & Rs!ChamsName)
        End If
        If Rs!CreatOnglet = True Then
            LstOngletName = Trim("" & Rs!ChamsName)
        End If
    End If

Next
If Trim(LstOngletName) <> "" Then
For L = 2 To MyRange.Rows.Count
Debug.Print MyRange(L, MyClonne(LstOngletName))
    If InStr(1, ";" & LstOnglet & ";", ";" & MyRange(L, MyClonne(LstOngletName)) & ";") = 0 Then LstOnglet = LstOnglet & ";" & MyRange(L, MyClonne(LstOngletName))
    Debug.Print LstOnglet
    
Next
End If
End Sub
Public Sub ScanDapresDoc(IndexTableauDoc As Long, ColecDoc As Collection, TableDoc As GenerateurDoc, SaveEn As String, MenuM1 As Boolean)
Dim copyIndexTableauDoc As Long
Dim Menu As String
Dim DapresDoc() As String
Dim I As Long
copyIndexTableauDoc = IndexTableauDoc
DapresDoc = TableDoc.DapresDocs
If indexdoc = 0 Then Exit Sub
For I = 0 To UBound(DapresDoc)
 Menu = DapresDoc(I)
  ChargementColecDoc "" & DapresDoc(I), ColecDoc, TableDoc, SaveEn, MenuM1
Next
End Sub
Public Function ChargemenDapresDoc(DocName As String, Mydoc As GenerateurDoc)
Dim I As Long
Dim tableau
ReDim Preserve DapresDoc(indexdoc)
If Trim(DocName) = "" Then GoTo Fin
If indexdoc = 0 Then
    
        DapresDoc(indexdoc) = DocName
        
        indexdoc = indexdoc + 1
Else
    For I = 0 To UBound(Mydoc.DapresDocs)
        If UCase(DapresDocs(I)) = UCase(DocName) Then GoTo Fin
    Next
    
    ReDim Preserve DapresDoc(indexdoc)
    DapresDoc(indexdoc) = DocName
    
    indexdoc = indexdoc + 1
End If
Fin:
ChargemenDapresDoc = DapresDoc
End Function
Private Sub ChargementColecDoc(Menu As String, ColecDoc As Collection, TableDoc, SaveEn As String, Optional MenuM1 As Boolean)
Dim Sql As String
Dim Rs As Recordset
Dim RsIdMenu As Recordset
Dim copyIndexTableauDoc As Long
Dim aa
Dim I As Long

Dim Mydoc As GenerateurDoc
Sql = "SELECT T_ETATS.ID, T_ETATS.EtatName, T_Etats_Onglet.Document, T_ETATS.Menu "
Sql = Sql & "FROM T_ETATS INNER JOIN T_Etats_Onglet ON T_ETATS.ID = T_Etats_Onglet.Id_Etat "
Sql = Sql & "WHERE T_ETATS.EtatName='" & MyReplace(Menu) & "' OR T_ETATS.Menu='" & MyReplace(Menu) & "' "
Sql = Sql & "GROUP BY T_ETATS.ID, T_ETATS.EtatName, T_Etats_Onglet.Document, T_ETATS.Menu;"

Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = True And Trim("" & Menu) <> "" Then

    On Error Resume Next
    aa = ""
    aa = ColecDoc("" & Menu).Menu
    If Err Then
        Err.Clear
 
        
        Set Mydoc = New GenerateurDoc
        ColecDoc.Add Mydoc, "" & Replace(Menu, " ", "_")
         Set Mydoc = Nothing
         End If
         GenerateurDocOk = True
         ColecDoc("" & Replace(Menu, " ", "_")).Menu = "" & Menu
         ColecDoc("" & Replace(Menu, " ", "_")).PathDoc = "" & Menu
         If MenuM1 = True Then
            LstColecDoc.Add "" & Replace(Menu, " ", "_"), Replace(Menu, " ", "_")
        End If
         If ColecDoc("" & Replace(Menu, " ", "_")).IndexTableauDoc = 0 Then
            ColecDoc("" & Replace(Menu, " ", "_")).SaveAs = SaveEn
        Else
            ColecDoc("" & Replace(Menu, " ", "_")).SaveAs = Environ("USERPROFILE") & "\Mes Documents\" & "" & Replace(Menu, " ", "_") & "_" & ColecDoc("" & Replace(Menu, " ", "_")).PathDoc
        End If
    
     On Error GoTo 0
End If
While Rs.EOF = False
On Error Resume Next
    aa = ""
    aa = ColecDoc("" & Replace("" & Rs!Menu, " ", "_")).Menu
    If Err Then
        Err.Clear
         
        
       
        
        Set Mydoc = New GenerateurDoc
        ColecDoc.Add Mydoc, "" & Replace("" & Rs!Menu, " ", "_")
        Set Mydoc = Nothing
        ColecDoc("" & Replace("" & Rs!Menu, " ", "_")).Menu = "" & Replace("" & Rs!Menu, " ", "_")
        aa = ColecDoc("" & Replace("" & Rs!Menu, " ", "_")).Menu
        
       ColecDoc("" & Replace("" & Rs!Menu, " ", "_")).PathDoc = "" & Rs!EtatName
       
        Set Mydoc = ColecDoc("" & Replace("" & Rs!Menu, " ", "_"))
        If MenuM1 = True Then
            LstColecDoc.Add "" & Replace("" & Rs!Menu, " ", "_"), Menu
        End If
      Mydoc.DapresDocs = ColecDoc("" & Replace("" & Rs!Menu, " ", "_")).ChargemenDapresDoc("" & Rs!Document, Mydoc)
       If ColecDoc("" & Replace("" & Rs!Menu, " ", "_")).IndexTableauDoc = 0 Then
            ColecDoc("" & Replace("" & Rs!Menu, " ", "_")).SaveAs = SaveEn
        Else
            ColecDoc("" & Replace("" & Rs!Menu, " ", "_")).SaveAs = Environ("USERPROFILE") & "\Mes Documents\" & "" & Replace("" & Rs!Menu, " ", "_") & "_" & ColecDoc("" & Replace("" & Rs!Menu, " ", "_")).PathDoc
        End If
       If Trim("" & Rs!Document) <> "" Then
       ColecDoc("" & Replace("" & Rs!Menu, " ", "_")).SurDoc = True
       End If
        Sql = "SELECT T_ETATS.ID From T_ETATS "
        Sql = Sql & "WHERE T_ETATS.EtatName='" & MyReplace("" & Rs!EtatName) & "';"
        Set RsIdMenu = Con.OpenRecordSet(Sql)
         ColecDoc("" & Replace("" & Rs!Menu, " ", "_")).id_Menu = RsIdMenu!Id
         Set RsIdMenu = Con.CloseRecordSet(RsIdMenu)
         Set Mydoc = ColecDoc("" & Replace("" & Rs!Menu, " ", "_"))
         copyIndexTableauDoc = IndexTableauDoc
          Mydoc.ScanDapresDoc copyIndexTableauDoc, ColecDoc, ColecDoc("" & Replace("" & Rs!Menu, " ", "_")), SaveEn, True
    Else
        Set Mydoc = ColecDoc("" & Replace("" & Rs!Menu, " ", "_"))
       ColecDoc("" & Replace("" & Rs!Menu, " ", "_")).DapresDocs = ColecDoc("" & Replace("" & Rs!Menu, " ", "_")).ChargemenDapresDoc("" & Rs!Document, ColecDoc("" & Replace("" & Rs!Menu, " ", "_")))
       If ColecDoc("" & Replace("" & Rs!Menu, " ", "_")).IndexTableauDoc = 0 Then
            ColecDoc("" & Replace("" & Rs!Menu, " ", "_")).SaveAs = SaveEn
        Else
            ColecDoc("" & Replace("" & Rs!Menu, " ", "_")).SaveAs = Environ("USERPROFILE") & "\Mes Documents\" & "" & Replace("" & Rs!Menu, " ", "_") & "_" & ColecDoc("" & Replace("" & Rs!Menu, " ", "_")).PathDoc
        End If
       If Trim("" & Rs!Document) <> "" Then
        ColecDoc("" & Replace("" & Rs!Menu, " ", "_")).SurDoc = True
        
       End If
        Sql = "SELECT T_ETATS.ID From T_ETATS "
        Sql = Sql & "WHERE T_ETATS.EtatName='" & Rs!EtatName & "';"
        Set RsIdMenu = Con.OpenRecordSet(Sql)
         ColecDoc("" & Replace("" & Rs!Menu, " ", "_")).id_Menu = RsIdMenu!Id
         Set RsIdMenu = Con.CloseRecordSet(RsIdMenu)
         Set Mydoc = ColecDoc("" & Replace("" & Rs!Menu, " ", "_"))
         copyIndexTableauDoc = IndexTableauDoc
          Mydoc.ScanDapresDoc copyIndexTableauDoc, ColecDoc, ColecDoc("" & Replace("" & Rs!Menu, " ", "_")), SaveEn, True
    End If
    On Error GoTo 0
    Rs.MoveNext
Wend

End Sub
Public Sub LoadColecDoc(IndexTableauDoc As Long, ColecDoc As Collection, TableDoc, SaveEn As String)
Dim I As Long
For I = 1 To ColecDoc.Count
    ChargementColecDoc ColecDoc(I).Menu, ColecDoc, TableDoc, SaveEn
Next

End Sub
Public Sub SelectEtat(IndexTableauDoc As Long, ColecDoc As Collection, Doc As GenerateurDoc, _
Id_IndiceProjet As Long, Fils As Long, SaveEn As String, NumDoc As Long)
Dim I As Long
Dim Sql As String
Dim Fso As New FileSystemObject
Dim Rs As Recordset
Dim Mydoc As GenerateurDoc
Dim a

If Doc.DocExecuter = True Then Exit Sub
On Error Resume Next
For I = UBound(DapresDocs) To 0 Step -1
If Err Then GoTo Sortir
        If DapresDocs(I) <> "" Then
            ColecDoc(LstColecDoc(Replace(Doc.DapresDocs(I), " ", "_"))).SelectEtat I, ColecDoc, ColecDoc(LstColecDoc(Replace(Doc.DapresDocs(I), " ", "_"))), Id_IndiceProjet, Fils, ColecDoc(LstColecDoc(Replace(Doc.DapresDocs(I), " ", "_"))).SaveAs, I
         End If
Next
Sortir:
Err.Clear
Select Case UCase(Doc.PathDoc)
    Case "LI"
         NomenclatureOk = True
        Sql = "SELECT T_Projet.Projet, T_indiceProjet.Vague, T_indiceProjet.Equipement, "
        Sql = Sql & "T_indiceProjet.Ensemble, [PI] & '_' & [PI_Indice] AS Piece, [PL] & '_' "
        Sql = Sql & "& [PL_Indice] AS Plan, [OU] & '_' & [OU_Indice] AS Outil, [Li] & '_' "
        Sql = Sql & "& [LI_Indice] AS Liste, T_indiceProjet.LI_Indice, T_indiceProjet.Client, "
        Sql = Sql & "[RefPieceClient] & '_' & [Ref_Piece_CLI] AS RfCli, T_indiceProjet.Id, T_indiceProjet.CleAc "
        Sql = Sql & "FROM T_Projet INNER JOIN (T_Pieces INNER JOIN T_indiceProjet "
        Sql = Sql & "ON T_Pieces.Id = T_indiceProjet.Id_Pieces) ON T_Projet.id = "
        Sql = Sql & "T_Pieces.IdProjet "
        Sql = Sql & "WHERE T_indiceProjet.Id=" & Id_IndiceProjet & ";"
        Set Rs = Con.OpenRecordSet(Sql)
        Set Mydoc = Doc
        
    If IndexTableauDoc <> 0 Then
   
        DocExecuter = ExporteXls(SaveEn & ".XLS", Id_IndiceProjet, Save:=False)
    End If
    Case "DOSSIER FAB"
        NomenclatureOk = True
        Sql = "SELECT T_Projet.Projet, T_indiceProjet.Vague, T_indiceProjet.Equipement, "
        Sql = Sql & "T_indiceProjet.Ensemble, [PI] & '_' & [PI_Indice] AS Piece, [PL] & '_' "
        Sql = Sql & "& [PL_Indice] AS Plan, [OU] & '_' & [OU_Indice] AS Outil, [Li] & '_' "
        Sql = Sql & "& [LI_Indice] AS Liste, T_indiceProjet.LI_Indice, T_indiceProjet.Client, "
        Sql = Sql & "[RefPieceClient] & '_' & [Ref_Piece_CLI] AS RfCli, T_indiceProjet.Id, T_indiceProjet.CleAc "
        Sql = Sql & "FROM T_Projet INNER JOIN (T_Pieces INNER JOIN T_indiceProjet "
        Sql = Sql & "ON T_Pieces.Id = T_indiceProjet.Id_Pieces) ON T_Projet.id = "
        Sql = Sql & "T_Pieces.IdProjet "
        Sql = Sql & "WHERE T_indiceProjet.Id=" & Id_IndiceProjet & ";"
        Set Rs = Con.OpenRecordSet(Sql)
        Set Mydoc = Doc
        
    If IndexTableauDoc - 1 <> 0 Then
   
        If ExporteXls(Environ("USERPROFILE") & "\Mes Documents\" & Doc.PathDoc & "LI.XLS", Id_IndiceProjet) = True Then
            EnteteClasseurControle = "Fabrication"
            
            DocExecuter = DossierDeFab(Environ("USERPROFILE") & "\Mes Documents\" & Doc.PathDoc & "LI.XLS" _
            , Id_IndiceProjet, "" & Rs!Projet, "" & Rs!Vague, "" & Rs!Equipement, "" & Rs!Ensemble, "" & Rs!Piece, "" & Rs!Plan, _
            "" & Rs!Outil, "" & Rs!Liste, "" & Rs!Client, "" & Rs!RfCli, "F_Encours", EnteteClasseurControle, "Nc", True, "" & Rs!CleAc, Id_IndiceProjet, SaveAs & ".XLS", False)
            
            If Fso.FileExists(Environ("USERPROFILE") & "\Mes Documents\" & Doc.PathDoc & "LI.XLS") = True Then Fso.DeleteFile Environ("USERPROFILE") & "\Mes Documents\" & Doc.PathDoc & "LI.XLS"
        End If
    End If
    Case UCase("Dossier Control")
         NomenclatureOk = True
        Sql = "SELECT T_Projet.Projet, T_indiceProjet.Vague, T_indiceProjet.Equipement, "
        Sql = Sql & "T_indiceProjet.Ensemble, [PI] & '_' & [PI_Indice] AS Piece, [PL] & '_' "
        Sql = Sql & "& [PL_Indice] AS Plan, [OU] & '_' & [OU_Indice] AS Outil, [Li] & '_' "
        Sql = Sql & "& [LI_Indice] AS Liste, T_indiceProjet.LI_Indice, T_indiceProjet.Client, "
        Sql = Sql & "[RefPieceClient] & '_' & [Ref_Piece_CLI] AS RfCli, T_indiceProjet.Id, T_indiceProjet.CleAc "
        Sql = Sql & "FROM T_Projet INNER JOIN (T_Pieces INNER JOIN T_indiceProjet "
        Sql = Sql & "ON T_Pieces.Id = T_indiceProjet.Id_Pieces) ON T_Projet.id = "
        Sql = Sql & "T_Pieces.IdProjet "
        Sql = Sql & "WHERE T_indiceProjet.Id=" & Id_IndiceProjet & ";"
        Set Rs = Con.OpenRecordSet(Sql)
        Set Mydoc = Doc
        
    If IndexTableauDoc <> 0 Then
   
        If ExporteXls(Environ("USERPROFILE") & "\Mes Documents\" & Doc.PathDoc & "LI.XLS", Id_IndiceProjet) = True Then
           EnteteClasseurControle = "Contrôle"
            
            DocExecuter = DossierDeFab(Environ("USERPROFILE") & "\Mes Documents\" & Doc.PathDoc & "LI.XLS" _
            , Id_IndiceProjet, "" & Rs!Projet, "" & Rs!Vague, "" & Rs!Equipement, "" & Rs!Ensemble, "" & Rs!Piece, "" & Rs!Plan, _
            "" & Rs!Outil, "" & Rs!Liste, "" & Rs!Client, "" & Rs!RfCli, "F_Encours", EnteteClasseurControle, "Nc", False, "" & Rs!CleAc, Id_IndiceProjet, SaveAs & ".XLS", False)
            
            If Fso.FileExists(Environ("USERPROFILE") & "\Mes Documents\" & Doc.PathDoc & "LI.XLS") = True Then Fso.DeleteFile Environ("USERPROFILE") & "\Mes Documents\" & Doc.PathDoc & "LI.XLS"
        End If
    End If
'
'        Case UCase("Nomenclature")
'         NomenclatureOk = False
'        Sql = "SELECT T_Projet.Projet, T_indiceProjet.Vague, T_indiceProjet.Equipement, "
'        Sql = Sql & "T_indiceProjet.Ensemble, [PI] & '_' & [PI_Indice] AS Piece, [PL] & '_' "
'        Sql = Sql & "& [PL_Indice] AS Plan, [OU] & '_' & [OU_Indice] AS Outil, [Li] & '_' "
'        Sql = Sql & "& [LI_Indice] AS Liste, T_indiceProjet.LI_Indice, T_indiceProjet.Client, "
'        Sql = Sql & "[RefPieceClient] & '_' & [Ref_Piece_CLI] AS RfCli, T_indiceProjet.Id, T_indiceProjet.CleAc "
'        Sql = Sql & "FROM T_Projet INNER JOIN (T_Pieces INNER JOIN T_indiceProjet "
'        Sql = Sql & "ON T_Pieces.Id = T_indiceProjet.Id_Pieces) ON T_Projet.id = "
'        Sql = Sql & "T_Pieces.IdProjet "
'        Sql = Sql & "WHERE T_indiceProjet.Id=" & Id_IndiceProjet & ";"
'        Set RS = Con.OpenRecordSet(Sql)
'        Set Mydoc = Doc
'
'    If IndexTableauDoc <> 0 Then
'
'        DocExecuter = ExporteXls(SaveEn & ".XLS", Id_IndiceProjet, Save:=False)
'    End If
    Case Else
'        DocExecuter = Module1.GenerEtat(id_Menu, Id_IndiceProjet, Fils, SaveEn, IsPDF, NumDoc)
'
'
'        If Doc.SurDoc = True Then
'
'           DocExecuter = GenerEtatDapresDoc(id_Menu, Id_IndiceProjet, Fils, SaveEn, IsPDF, Doc, ColecDoc, NumDoc - 1)
'        Else
''
                DocExecuter = GenerEtats(id_Menu, Id_IndiceProjet, Fils, SaveEn, IsPDF, ColecDoc, SaveEn, NumDoc, Doc.SurDoc)
'
'        End If
'
End Select
Set Fso = Nothing
End Sub

Private Function BoucleOnglets(Id_Onglet As Long, MyWorkbookSource As Workbook)
Dim IndexPlages As PlageOnget
Dim Sql As String
Dim RsOptionOnglet As Recordset
Dim DebutOnglet, FinOnget, I As Long

Sql = "SELECT T_Etats_Onglet.* From T_Etats_Onglet WHERE T_Etats_Onglet.Id=" & Id_Onglet & ";"
Set RsOptionOnglet = Con.OpenRecordSet(Sql)
If Trim("" & RsOptionOnglet!OngletStrat) <> "" And Trim("" & RsOptionOnglet!OngleEnd) <> "" And RsOptionOnglet!FiltreSequentielle = True Then
    If UCase(Trim("" & RsOptionOnglet!OngletStrat)) = "PREMIER" Then
        DebutOnglet = 1
    Else
        If IsNumeric(Trim("" & RsOptionOnglet!OngletStrat)) Then
            DebutOnglet = Val(Trim("" & RsOptionOnglet!OngletStrat))
        Else
            For I = 1 To MyWorkbookSource.Sheets.Count
                If UCase(MyWorkbookSource.Sheets(I).Name) = UCase(Trim("" & RsOptionOnglet!OngletStrat)) Then
                    DebutOnglet = I
                    Exit For
                End If
            Next
            If DebutOnglet = 0 Then
                MyPlageOnglets.Err = True
                Exit Function
             End If
        End If
    End If
        DebutOnglet = DebutOnglet + Val("" & RsOptionOnglet!DecaleAppres)
        If DebutOnglet < 1 And DebutOnglet > MyWorkbookSource.Sheets.Count Then
            MyPlageOnglets.Err = True
            Exit Function
        End If

        If UCase(Trim("" & RsOptionOnglet!OngleEnd)) = "DERNIER" Then
            FinOnget = MyWorkbookSource.Sheets.Count
        Else
            If IsNumeric(Trim("" & RsOptionOnglet!OngleEnd)) Then
                FinOnget = Val(Trim("" & RsOptionOnglet!OngleEnd))
            Else
                For I = 1 To MyWorkbookSource.Sheets.Count
                    If UCase(MyWorkbookSource.Sheets(I).Name) = UCase(Trim("" & RsOptionOnglet!OngleEnd)) Then
                        FinOnget = I
                        Exit For
                    End If
                Next
                If FinOnget = 0 Then
                    MyPlageOnglets.Err = True
                    Exit Function
                End If
            End If
        End If
        FinOnget = FinOnget - Val("" & RsOptionOnglet!DecaleAvant)
        If FinOnget < 1 And FinOnget > MyWorkbookSource.Sheets.Count Then
            MyPlageOnglets.Err = True
            Exit Function
        End If
        If DebutOnglet > FinOnget Then
            MyPlageOnglets.Err = True
        Else
            MyPlageOnglets.Err = False
            MyPlageOnglets.Debut = DebutOnglet
            MyPlageOnglets.Fin = FinOnget
        End If
Else
If Trim("" & RsOptionOnglet!Document) <> "" Then

     For I = 1 To MyWorkbookSource.Sheets.Count
      Debug.Print UCase(MyWorkbookSource.Sheets(I).Name)
                If UCase(MyWorkbookSource.Sheets(I).Name) = UCase(Trim("" & RsOptionOnglet!OngletStrat)) Then
                    DebutOnglet = I
                    Exit For
                End If
            Next
            If DebutOnglet = 0 Then
                MyPlageOnglets.Err = True
                Exit Function
             End If
             If DebutOnglet > MyWorkbookSource.Sheets.Count Then
              MyPlageOnglets.Err = True
                Exit Function
             End If
             MyPlageOnglets.Err = False
            MyPlageOnglets.Debut = DebutOnglet
            MyPlageOnglets.Fin = DebutOnglet
 Else
            MyPlageOnglets.Err = False
            MyPlageOnglets.Debut = 1
            MyPlageOnglets.Fin = 1
 End If
End If
End Function
Private Function GenerEtats(Id_Macro As Long, Id_IndiceProjet As Long, Fils As Long, _
                            FileExcel As String, PDF As Boolean, ColecDoc As Collection, SaveEn As String, NumDoc As Long, SurDoc As Boolean) As Boolean
                            
Dim Trie As String
Dim I As Long
Dim ChoixTraitement As Long
Dim Sql As String
Dim RsMacro As Recordset
Dim RsSelectChamp As Recordset
Dim RsExcel As Recordset
Dim Table As String
Dim MyWorkbook As EXCEL.Workbook
Dim MySheet As Worksheet
Set MyExcel = EXCEL.Application
Dim SqlChamp As String
Dim Fso As New FileSystemObject
Dim FicherEntrer As String
Sql = "SELECT DISTINCT  T_ETATS.EtatName,T_ETATS.ID, T_Etats_Onglet.Id as Id_Select, "
Sql = Sql & "T_Etats_Onglet.Onglet, T_Etats_Onglet.Document ,T_Etats_Onglet.SaveOnglet, T_Menu_Etat_Onglet.TableName,T_Etats_Onglet.VueEpissur, T_Etats_Onglet.GestionOnglet, T_Etats_Onglet.GestionOnglet2,T_Etats_Onglet.PerfEntete,T_Etats_Onglet.Paysage "
Sql = Sql & "FROM (T_ETATS INNER JOIN T_Etats_Onglet ON T_ETATS.ID = T_Etats_Onglet.Id_Etat) INNER JOIN T_Menu_Etat_Onglet ON T_Etats_Onglet.Onglet = T_Menu_Etat_Onglet.Menu "
Sql = Sql & "Where T_ETATS.Id = " & Id_Macro & " "
Sql = Sql & "ORDER BY T_Etats_Onglet.Id;"

BoolGenEtatEpisure = False
If SurDoc = True Then ChoixTraitement = 1
    
Set RsMacro = Con.OpenRecordSet(Sql)

While RsMacro.EOF = False
Set MyExcel = New EXCEL.Application
MyExcel.Visible = True
MyExcel.DisplayAlerts = False
If DocExecuter = False Then
    Set MyWorkbook = MyExcel.Workbooks.Add
Else
     Set MyWorkbook = MyExcel.Workbooks.Open(SaveAs)
End If

For I = 1 To MyWorkbook.Worksheets.Count
    MyWorkbook.Worksheets(I).Name = Replace(MyWorkbook.Worksheets(I).Name, "Feuil", "§Feuil§", 1)
Next

MyExcel.Visible = True
 MyExcel.DisplayAlerts = False
GestionOnglets = RsMacro!GestionOnglet
GestionOnglets2 = RsMacro!GestionOnglet2
PerfEntete = RsMacro!PerfEntete
If RsMacro!Paysage = True Then
    PortraitPaysage = xlLandscape
Else
    PortraitPaysage = xlPortrait
End If


If NumDoc = 0 Then bool_MiseEnPage = True
ComteurPass = ComteurPass + 1
Debug.Print "ComteurPass : " & ComteurPass
    If RsMacro!VueEpissur = True Then BoolGenEtatEpisure = True
        Select Case ChoixTraitement
            Case 0
                GenerEtats = GenerEtat(MyWorkbook, RsMacro!TableName, RsMacro!Id_Select, Id_Macro, Id_IndiceProjet, Fils, FileExcel, PDF, NumDoc, "" & RsMacro!SaveOnglet)
            Case 1
                 DocExecuter = GenerEtatSurDoc(MyExcel, MyWorkbook, "" & RsMacro!Document, RsMacro!TableName, RsMacro!Id_Select, Id_Macro, Id_IndiceProjet, Fils, FileExcel, PDF, NumDoc, "" & RsMacro!SaveOnglet, ColecDoc)
        End Select
        If NumDoc = 0 Then
             StopOnglet = StarOnglet + OfsetOnglet
             
            Set MyWorkbook = MyExcel.Workbooks(MyWorkbook.Name)
            DeletSheetEtat MyWorkbook, GestionOnglets, NumDoc, StarOnglet, StopOnglet
            Set MyWorkbook = MyExcel.Workbooks(MyWorkbook.Name)
            DeletSheetEtat MyWorkbook, GestionOnglets2, NumDoc, StarOnglet, StopOnglet

            StopOnglet = 0: StarOnglet = 0
'            If Fso.FileExists(SaveEn & ".XLS") Then
'            Fso.DeleteFile SaveEn & ".XLS"
'            End If
                        MyWorkbook.SaveAs SaveEn
                        DocExecuter = True
        Else
'        If Fso.FileExists(Environ("USERPROFILE") & "\Mes Documents\" & "" & Replace("" & Menu, " ", "_") & "_" & ColecDoc("" & Replace("" & Menu, " ", "_")).PathDoc & ".XLS") Then
'            Fso.DeleteFile Environ("USERPROFILE") & "\Mes Documents\" & "" & Replace("" & Menu, " ", "_") & "_" & ColecDoc("" & Replace("" & Menu, " ", "_")).PathDoc & ".XLS"
'        End If
'
            MyWorkbook.SaveAs Environ("USERPROFILE") & "\Mes Documents\" & "" & Replace("" & Menu, " ", "_") & "_" & ColecDoc("" & Replace("" & Menu, " ", "_")).PathDoc
            
        End If
        DoEvents
        MyExcel.Quit
        Set MyExcel = Nothing
         bool_MiseEnPage = False
    RsMacro.MoveNext
Wend
 MyWorkbook.Close False

End Function
Private Function InitXls(MyExcel As EXCEL.Application, MyWorkbook As EXCEL.Workbook, Document As String, Table As String, Id_Select As Long, Id_Macro As Long, Id_IndiceProjet As Long, Fils As Long, FileExcel As String, PDF As Boolean, NbDoc As Long, SaveOnglet As String, ColecDoc As Collection) As Workbook
Dim Sql As String
Dim RsDoc As Recordset
Dim RsExcel As Recordset
Dim MyWorkbookSource As Workbook
Dim MySheet As Worksheet
Dim OnGletName As String
Dim MySeet As Worksheet
Dim MyRange As Range
Dim I As Long
For I = NbDoc + 2 To ColecDoc.Count
   If ColecDoc(I).Menu = Document Then
    Set InitXls = MyExcel.Workbooks.Open(ColecDoc(I).SaveAs)
    Exit Function
   End If
Next
Sql = "SELECT T_Etats_Onglet.* FROM T_Etats_Onglet "
Sql = Sql & "WHERE T_Etats_Onglet.Id=" & Id_Select & ";"
Set RsDoc = Con.OpenRecordSet(Sql)
If Trim("" & RsDoc!Document) <> "" Then
    
    Sql = "SELECT T_ETATS.* FROM T_ETATS "
    Sql = Sql & "WHERE T_ETATS.EtatName='" & Replace(Document, "'", "''") & "';"
     Set RsDoc = Con.OpenRecordSet(Sql)
     If RsDoc.EOF = False Then
     Set MyWorkbookSource = MyExcel.Workbooks.Open(Environ("USERPROFILE") & "\Mes Documents\" & "" & Replace("" & RsDoc!Menu, " ", "_") & "_" & Replace("" & RsDoc!EtatName, " ", "_") & ".XLS")
     Else
      Set MyWorkbookSource = MyExcel.Workbooks.Open(ColecDoc(Document).SaveAs)
     End If
Else
    OnGletName = "" & RsDoc!SaveOnglet
        Sql = "select [" & Table & "].*  FROM " & Table & " Where " & Table & ".Id_IndiceProjet=" & Id_IndiceProjet & " "
        Sql = Sql & "ORDER BY [" & Table & "].ID;"
        Set RsExcel = Con.OpenRecordSet(Sql)
        Set MyWorkbookSource = MyExcel.Workbooks.Add
       Set MySeet = IsertSheet(MyWorkbookSource, OnGletName)
        Set MyRange = MySeet.Range("A1").CurrentRegion
    ExcelCreatTitre MyRange, RsExcel


 Sql = "SELECT DISTINCT T_Etats_Onglet.Id AS Id_Select, T_Etats_Select_Champs.ChamsName, T_Etats_Select_Champs.Visible "
    Sql = Sql & "FROM T_Etats_Onglet INNER JOIN T_Etats_Select_Champs ON T_Etats_Onglet.Id = T_Etats_Select_Champs.Id_Onglet "
    Sql = Sql & "Where T_Etats_Onglet.Id = " & Id_Select & " "
    Sql = Sql & "And T_Etats_Select_Champs.Visible = false; "
    Set RsChampMasquer = Con.OpenRecordSet(Sql)
If RsExcel.EOF = False Then
MySeet.Range("A2").CopyFromRecordset RsExcel
      ReplaceBool MySeet, "A2:" & MySeet.Range(Replace(MySeet.Cells(MySeet.Range("A1").CurrentRegion.Rows.Count, MySeet.Range("A1").CurrentRegion.Columns.Count).Address, "$", ""))
End If
'While RsExcel.EOF = False
'    Set Myrange = MySeet.Range("A1").CurrentRegion
'    Set Myrange = MySeet.Range(Myrange(Myrange.Rows.Count + 1, 1).Address & ":" & Myrange(Myrange.Rows.Count + 1, Myrange.Columns.Count).Address)
'    ExcelCreatTitre Myrange, RsExcel, True, Formule:=True
'    RsExcel.MoveNext
' Wend

RenseigneColonn MySeet, Id_Select, 1
        
End If



Set InitXls = MyWorkbookSource
Set RsDoc = Con.CloseRecordSet(RsDoc)
End Function

Function GenerEtatSurDoc(MyExcel As EXCEL.Application, MyWorkbook As EXCEL.Workbook, Document As String, Table As String, Id_Select As Long, Id_Macro As Long, Id_IndiceProjet As Long, Fils As Long, FileExcel As String, PDF As Boolean, NbDoc As Long, SaveOnglet As String, ColecDoc As Collection) As Boolean
Dim Trie As String
Dim I As Long
Dim Sql As String
Dim RsMacro As Recordset
Dim RsSelectChamp As Recordset
Dim RsExcel As Recordset
Dim MySheet As Worksheet
Dim SqlChamp As String
Dim RsDoc As Recordset

Dim MyWorkbookSource As Workbook

Set MyWorkbookSource = InitXls(MyExcel, MyWorkbook, Document, Table, Id_Select, Id_Macro, Id_IndiceProjet, Fils, FileExcel, PDF, NbDoc, SaveOnglet, ColecDoc)

Sql = "SELECT T_Etats_Onglet.Macro FROM T_Etats_Onglet "
Sql = Sql & "WHERE T_Etats_Onglet.Id=" & Id_Select & ";"
Set RsSelectChamp = Con.OpenRecordSet(Sql)
GeneEtatMacro = "" & RsSelectChamp!Macro
    Sql = "SELECT T_Etats_Select_Champs.ChamsName, T_Etats_Select_Champs.ChampAs, T_Etats_Select_Champs.Trie,T_Etats_Select_Champs.Visible "
Sql = Sql & "From T_Etats_Select_Champs "
Sql = Sql & "Where T_Etats_Select_Champs.Id_Onglet =" & Id_Select & " "
Sql = Sql & "ORDER BY T_Etats_Select_Champs.Id;"

    
    Set RsSelectChamp = Con.OpenRecordSet(Sql)
    If RsSelectChamp.EOF = False Then
        
            Sql = "Select "
            Trie = ""
        While RsSelectChamp.EOF = False
        If RsSelectChamp!Trie <> 0 Then
            Trie = Trie & "§" & RsSelectChamp!ChamsName & "=" & RsSelectChamp!Trie
        End If
            Sql = Sql & "[" & Table & "].[" & RsSelectChamp!ChamsName & "]"
            If UCase(RsSelectChamp!ChamsName) <> UCase(RsSelectChamp!ChampAs) Then
            Sql = Sql & " AS [" & RsSelectChamp!ChampAs & "]"
            End If
            Sql = Sql & ","
            RsSelectChamp.MoveNext
        Wend
        Sql = Left(Sql, Len(Sql) - 1)
        Sql = Sql & " FROM " & Table & " Where " & Table & ".Id_IndiceProjet=" & Id_IndiceProjet & "; "
        
        SqlChamp = Sql
        Sql = "select [" & Table & "].*  FROM " & Table & " Where " & Table & ".Id_IndiceProjet=" & Id_IndiceProjet & " "
        Sql = Sql & "ORDER BY [" & Table & "].ID;"
        Set RsExcel = Con.OpenRecordSet(Sql)
        
        
        If NbDoc = 0 Then
            CreerXls2DapresDoc MyExcel, MyWorkbook, MyWorkbookSource, RsExcel, SaveOnglet, Trie, Id_Select, Id_IndiceProjet, Fils, Id_Macro, NbDoc
        Else
            CreerXls2DapresDoc MyExcel, MyWorkbook, MyWorkbookSource, RsExcel, SaveOnglet, Trie, Id_Select, Id_IndiceProjet, Fils, Id_Macro, NbDoc, True
        End If
        Set RsExcel = Con.CloseRecordSet(RsExcel)
     End If
     
    Set RsSelectChamp = Con.CloseRecordSet(RsSelectChamp)
GenerEtatSurDoc = True

End Function
Function GenerEtat(MyWorkbook As EXCEL.Workbook, Table As String, Id_Select As Long, Id_Macro As Long, Id_IndiceProjet As Long, Fils As Long, FileExcel As String, PDF As Boolean, NbDoc As Long, SaveOnglet As String) As Boolean
Dim Trie As String
Dim I As Long
Dim Sql As String
Dim RsMacro As Recordset
Dim RsSelectChamp As Recordset
Dim RsExcel As Recordset
Dim MySheet As Worksheet
Dim SqlChamp As String
StarOnglet = MyWorkbook.Sheets.Count + 1
OfsetOnglet = -1
    Sql = "SELECT T_Etats_Select_Champs.ChamsName, T_Etats_Select_Champs.ChampAs, T_Etats_Select_Champs.Trie,T_Etats_Select_Champs.Visible "
Sql = Sql & "From T_Etats_Select_Champs "
Sql = Sql & "Where T_Etats_Select_Champs.Id_Onglet =" & Id_Select & " "
Sql = Sql & "ORDER BY T_Etats_Select_Champs.Id;"

    
    Set RsSelectChamp = Con.OpenRecordSet(Sql)
    If RsSelectChamp.EOF = False Then
        
            Sql = "Select "
            Trie = ""
        While RsSelectChamp.EOF = False
        If RsSelectChamp!Trie <> 0 Then
            Trie = Trie & "§" & RsSelectChamp!ChamsName & "=" & RsSelectChamp!Trie
        End If
            Sql = Sql & "[" & Table & "].[" & RsSelectChamp!ChamsName & "]"
            If UCase(RsSelectChamp!ChamsName) <> UCase(RsSelectChamp!ChampAs) Then
            Sql = Sql & " AS [" & RsSelectChamp!ChampAs & "]"
            End If
            Sql = Sql & ","
            RsSelectChamp.MoveNext
        Wend
        Sql = Left(Sql, Len(Sql) - 1)
        Sql = Sql & " FROM " & Table & " Where " & Table & ".Id_IndiceProjet=" & Id_IndiceProjet & "; "
        
        SqlChamp = Sql
        Sql = "select [" & Table & "].*  FROM " & Table & " Where " & Table & ".Id_IndiceProjet=" & Id_IndiceProjet & " "
        Sql = Sql & "ORDER BY [" & Table & "].ID;"
        Set RsExcel = Con.OpenRecordSet(Sql)
        
        
        If NbDoc = 0 Then
       bool_MiseEnPage = True

            CreerXls2 MyExcel, MyWorkbook, RsExcel, SaveOnglet, Trie, Id_Select, Id_IndiceProjet, Fils, Id_Macro, NbDoc
        Else
            CreerXls2 MyExcel, MyWorkbook, RsExcel, SaveOnglet, Trie, Id_Select, Id_IndiceProjet, Fils, Id_Macro, NbDoc, True
        End If
        Set RsExcel = Con.CloseRecordSet(RsExcel)
     End If
      bool_MiseEnPage = False
    Set RsSelectChamp = Con.CloseRecordSet(RsSelectChamp)
GenerEtat = True
End Function
Private Sub Class_Initialize()
IndexTableauDoc = IndexTableauDocGen

IndexTableauDocGen = IndexTableauDocGen + 1
End Sub

Function CreatEtaExels(MyExel As EXCEL.Workbook, Rs As Recordset, Onglet As String, Trie As String, _
                        Id_Onglet As Long, Id_IndiceProjet As Long, Fils As Long, Id_Macro As Long, _
                        NbDoc As Long, Optional DapresDoc As Boolean) As Worksheet
Dim MySeet As EXCEL.Worksheet
Dim MyRange As EXCEL.Range
Dim SheetFiltre As Worksheet
Dim SheetFiltreCible As Worksheet
Dim SaveNbL As Long
Set MySeet = IsertSheet(MyExel, Onglet)
Set MyRange = MySeet.Range("A1").CurrentRegion
Dim RsSelectChamp As Recordset
Dim RsCriteres As Recordset
Dim MyTrie
Dim MyTrieS
Dim Sql As String
Dim RsFiltre As Recordset
Dim RsFiltreELab As Recordset
Dim I As Long
Dim TrieExcel(3, 1) As String
Dim Col As Long
Dim NbTrie As Integer
Dim IsFilterElab As Boolean
Dim MyOption
Dim MyOption2
Dim NbMyOption As Long
Dim RangeCopy As Range
Dim PoseFiterMyOption As Long
Dim Debut As Long
Dim I2, I3, I4, I5 As Long
Dim L, C As Long
Dim OpTionValue
Dim RsFiterSurVue As Recordset
Dim ClsDog As GenerateurDoc
Dim SplitLstOnglet
'Dim RsChampMasquer As Recordset
Dim KillColImPaire As Boolean
Dim PasFiltre As Boolean
Dim PrefixName As String
Dim ChoixTraitement As Long
PrefixName = "" ' "_FLT_S_Equ_"

ExcelCreatTitre MyRange, Rs


 Sql = "SELECT DISTINCT T_Etats_Onglet.Id AS Id_Select, T_Etats_Select_Champs.ChamsName, T_Etats_Select_Champs.Visible "
    Sql = Sql & "FROM T_Etats_Onglet INNER JOIN T_Etats_Select_Champs ON T_Etats_Onglet.Id = T_Etats_Select_Champs.Id_Onglet "
    Sql = Sql & "Where T_Etats_Onglet.Id = " & Id_Onglet & " "
    Sql = Sql & "And T_Etats_Select_Champs.Visible = false; "
    Set RsChampMasquer = Con.OpenRecordSet(Sql)

While Rs.EOF = False
    Set MyRange = MySeet.Range("A1").CurrentRegion
    Set MyRange = MySeet.Range(MyRange(MyRange.Rows.Count + 1, 1).Address & ":" & MyRange(MyRange.Rows.Count + 1, MyRange.Columns.Count).Address)
    ExcelCreatTitre MyRange, Rs, True, Formule:=True
    Rs.MoveNext
 Wend

RenseigneColonn MySeet, Id_Onglet, 1
 MyTrie = Split(Trie & "§", "§")
 If UBound(MyTrie) > 1 Then
    For I = 1 To UBound(MyTrie)
        MyTrieS = Split(MyTrie(I) & "=", "=")
    If Val(MyTrieS(1)) <> 0 Then
    TrieExcel(I, 0) = MyTrieS(0): TrieExcel(I, 1) = MyTrieS(1)
    If I = 3 Then Exit For
    End If
   
 Next
 Set MyRange = MySeet.Range("A1").CurrentRegion
NbTrie = 0
 For I = 1 To 3
    For Col = 1 To MyRange.Columns.Count
       If MyRange(1, Col) = TrieExcel(I, 0) Then
       NbTrie = NbTrie + 1
           TrieExcel(I, 0) = MyRange(2, Col).Address
            TrieExcel(I, 0) = Replace(TrieExcel(I, 0), "$", "")
           Exit For
       End If
    Next
  Next
 End If
 Trier MySeet, NbTrie, Replace("A2:" & MyRange(MyRange.Rows.Count, MyRange.Columns.Count).Address, "$", ""), TrieExcel(1, 0), Val(TrieExcel(1, 1)), TrieExcel(2, 0), Val(TrieExcel(2, 1)), TrieExcel(3, 0), Val(TrieExcel(3, 1))
 Set MyRange = MySeet.Range("A1").CurrentRegion
 For I = 1 To MyRange.Columns.Count
    If UCase(MyRange(1, I)) = "OPTION" Or UCase(MyRange(1, I)) = UCase("Criteres") Then
        For I2 = 2 To MyRange.Rows.Count
            MyRange(I2, I) = ";" & MyRange(I2, I) & ";"
            
        Next
        Exit For
    End If
 Next
 
 Set CollecFiltreName = Nothing
 Set CollecFiltreName = New Collection
 Sql = "SELECT T_Etats_Select_Filtre.FiltreName "
Sql = Sql & "From T_Etats_Select_Filtre "
Sql = Sql & "Where T_Etats_Select_Filtre.Id_Onglet =  " & Id_Onglet & " "
Sql = Sql & "GROUP BY T_Etats_Select_Filtre.FiltreName "
Sql = Sql & "ORDER BY T_Etats_Select_Filtre.FiltreName; "

Set RsFiltre = Con.OpenRecordSet(Sql)
I = 0
If RsFiltre.EOF = True Then
    
Sql = "SELECT T_Etats_Select_Champs.ChampAs, T_indiceProjet.Equipement, T_Etats_Select_Champs.Id_Onglet "
        Sql = Sql & "FROM T_indiceProjet, T_Etats_Onglet INNER JOIN T_Etats_Select_Champs ON T_Etats_Onglet.Id = T_Etats_Select_Champs.Id_Onglet "
        Sql = Sql & "WHERE (T_Etats_Select_Champs.ChamsName='OPTION' OR T_Etats_Select_Champs.ChamsName='CRITERES') "
        If Fils = 0 Then
         Sql = Sql & "AND T_indiceProjet.Id=" & Id_IndiceProjet & " "
        Else
         Sql = Sql & "AND T_indiceProjet.Id=" & Fils & " "
        End If
        Sql = Sql & "AND T_Etats_Select_Champs.Id_Onglet=" & Id_Onglet & " AND T_Etats_Onglet.FiltreEquipement=True "
        Sql = Sql & "GROUP BY T_Etats_Select_Champs.ChamsName, T_Etats_Select_Champs.ChampAs, T_indiceProjet.Equipement, T_indiceProjet.Id, "
        Sql = Sql & "T_Etats_Select_Champs.Id_Onglet;"
        

Set RsFiltreELab = Con.OpenRecordSet(Sql)
If RsFiltreELab.EOF = False Then
IsFilterElab = True
MyOption = FunCodeCritaire(Id_IndiceProjet, "" & RsFiltreELab!Equipement)
 Set SheetFiltre = IsertSheet(MyExel, "MyFilterElaborré")
    DeleteRow SheetFiltre, True
    ExcelCreatTitre SheetFiltre.Range("a1").CurrentRegion, Rs
   Sql = "SELECT T_Etats_Select_Champs.ChampAs, T_Etats_Select_Champs.CreatOnglet "
    Sql = Sql & "FROM T_Etats_Onglet INNER JOIN T_Etats_Select_Champs ON T_Etats_Onglet.Id = T_Etats_Select_Champs.Id_Onglet  "
    Sql = Sql & "Where T_Etats_Select_Champs.Id_Onglet = " & RsFiltreELab!Id_Onglet & "  "
    Sql = Sql & " And T_Etats_Select_Champs.CreatOnglet=True "
    Sql = Sql & "GROUP BY T_Etats_Select_Champs.ChampAs, T_Etats_Select_Champs.CreatOnglet, T_Etats_Select_Champs.ChamsName;"
        
            For I2 = 0 To UBound(MyOption)
            For I3 = 1 To SheetFiltre.Range("a1").CurrentRegion.Columns.Count
                If SheetFiltre.Cells(1, I3) = RsFiltreELab!ChampAs Then
                PrefixName = "_FLT_S_Equ_"
                     SheetFiltre.Cells(I2 + 2, I3) = "*;" & UCase(Replace("" & MyOption(I2), "§Null§", "") & ";*")
                  
                   
                    Exit For
                 End If
                
            Next
         Next
      
    
     Set MyRange = MySeet.Range("a1").CurrentRegion
     For I2 = 1 To MyRange.Columns.Count
            If MyRange(1, I2) = "OPTION" Then Exit For
         Next
          For I3 = 2 To MyRange.Rows.Count
            If Trim("" & MyRange(I3, I2)) = "" Then MyRange(I3, I2) = ";;"
        Next
    End If
   
   If RsFiltreELab.EOF = False Then ChoixTraitement = 1
    If Trim(LstOnglet) <> "" Then ChoixTraitement = ChoixTraitement + 2
    If Trim(LstOnglet) = "" Then
    Set SheetFiltre = IsertSheet(MyExel, "MyFilterElaborré")
            If SheetFiltre.Range("A1").CurrentRegion.Rows.Count = 1 And SheetFiltre.Range("A1").CurrentRegion.Columns.Count = 1 Then
                If Trim("" & SheetFiltre.Range("A1")) = "" Then
                    
                    For C = 1 To MySeet.Range("a1").CurrentRegion.Columns.Count
                        SheetFiltre.Cells(1, C) = MySeet.Cells(1, C)
                        
                    Next
                End If
             End If
        If MySeet.Name <> MySeet.Name & PrefixName Then
            Set SheetFiltreCible = IsertSheet(MyExel, MySeet.Name & PrefixName, True)
            FiltreActif MySeet.Range("a1").CurrentRegion, SheetFiltre.Range("A1").CurrentRegion, SheetFiltreCible.Range("a1").CurrentRegion
         Else
             Set SheetFiltreCible = MySeet
        End If
      While RsChampMasquer.EOF = False

                    SheetFiltreCible.Range(Replace(SheetFiltreCible.Cells(1, MyClonne("" & RsChampMasquer!ChamsName)).Address, "$", "")).Interior.ColorIndex = 6
                     RsChampMasquer.MoveNext
                Wend
                If DapresDoc = False Then
                    For I5 = SheetFiltreCible.Range("A1").CurrentRegion.Columns.Count To 1 Step -1
                        If SheetFiltreCible.Cells(1, I5).Interior.ColorIndex <> 6 Then
                            SuprmerCells SheetFiltreCible.Range(Replace(SheetFiltreCible.Cells(1, I5).Address & ":" & SheetFiltreCible.Cells(SheetFiltreCible.Range("A1").CurrentRegion.Rows.Count, I5).Address, "$", "")), "g"
                            If KillColImPaire = True Then
                                SuprmerCells SheetFiltreCible.Range(Replace(SheetFiltreCible.Cells(SheetFiltreCible.Range("A1").CurrentRegion.Rows.Count + 1, 1).Address & ":" & SheetFiltreCible.Cells(SheetFiltreCible.Range("A1").CurrentRegion.Rows.Count + 1000, 1).Address, "$", "")), "g"
                                KillColImPaire = False
                            Else
                                KillColImPaire = True
                            End If
                         End If
                     Next
                End If
        If CrerSoutotal(SheetFiltreCible, Id_Onglet, Id_IndiceProjet, Fils, NbDoc) = False Then
            DeletSheet SheetFiltreCible
        End If
    Else
    SplitLstOnglet = Split(LstOnglet & ";", ";")
    For I = 2 To MySeet.Range("a1").CurrentRegion.Rows.Count
        MySeet.Cells(I, MyClonne(LstOngletName)) = ";" & MySeet.Cells(I, MyClonne(LstOngletName)) & ";"
    Next
            For I4 = 0 To UBound(SplitLstOnglet)
            If Trim(SplitLstOnglet(I4)) <> "" Then
            Set SheetFiltre = IsertSheet(MyExel, "MyFilterElaborré")
            If SheetFiltre.Range("A1").CurrentRegion.Rows.Count = 1 And SheetFiltre.Range("A1").CurrentRegion.Columns.Count = 1 Then
                If Trim("" & SheetFiltre.Range("A1")) = "" Then
                    PrefixName = "_Ongt_"
                    For C = 1 To MySeet.Range("a1").CurrentRegion.Columns.Count
                        SheetFiltre.Cells(1, C) = MySeet.Cells(1, C)
                        SheetFiltre.Cells(2, MyClonne(LstOngletName)) = "a"
                    Next
                End If
            End If

                For I = 2 To SheetFiltre.Range("a1").CurrentRegion.Rows.Count
                    SheetFiltre.Cells(I, MyClonne(LstOngletName)) = ";" & SplitLstOnglet(I4) & ";"
                Next
                  
                        Set SheetFiltreCible = IsertSheet(MyExel, MySeet.Name & PrefixName & SplitLstOnglet(I4), True)
                 
             
                FiltreActif MySeet.Range("a1").CurrentRegion, SheetFiltre.Range("A1").CurrentRegion, SheetFiltreCible.Range("a1").CurrentRegion
               OfseteMiseEnPage = SurEquipement(PrefixName, MyWorkbook, SheetFiltreCible, Id_IndiceProjet, Id_Onglet, Fils)
                For I = 2 To SheetFiltreCible.Range("a1").CurrentRegion.Rows.Count
                    SheetFiltreCible.Cells(I, MyClonne(LstOngletName)) = Replace(SheetFiltreCible.Cells(I, MyClonne(LstOngletName)), ";", "")
                Next
                
                    RsChampMasquer.Requery
                    KillColImPaire = False
                    SheetFiltreCible.Cells.Replace ";", ""
                While RsChampMasquer.EOF = False

                    SheetFiltreCible.Range(Replace(SheetFiltreCible.Cells(1, MyClonne("" & RsChampMasquer!ChamsName)).Address, "$", "")).Interior.ColorIndex = 6
                     RsChampMasquer.MoveNext
                Wend
                If DapresDoc = False Then
                    For I5 = SheetFiltreCible.Range("A1").CurrentRegion.Columns.Count To 1 Step -1
                        If SheetFiltreCible.Cells(1, I5).Interior.ColorIndex <> 6 Then
                            SuprmerCells SheetFiltreCible.Range(Replace(SheetFiltreCible.Cells(1, I5).Address & ":" & SheetFiltreCible.Cells(SheetFiltreCible.Range("A1").CurrentRegion.Rows.Count, I5).Address, "$", "")), "g"
                            If KillColImPaire = True Then
                                SuprmerCells SheetFiltreCible.Range(Replace(SheetFiltreCible.Cells(SheetFiltreCible.Range("A1").CurrentRegion.Rows.Count + 1, 1).Address & ":" & SheetFiltreCible.Cells(SheetFiltreCible.Range("A1").CurrentRegion.Rows.Count + 1000, 1).Address, "$", "")), "g"
                                KillColImPaire = False
                            Else
                                KillColImPaire = True
                            End If
                         End If
                     Next
                End If
               If DapresDoc = False Then
                If CrerSoutotal(SheetFiltreCible, Id_Onglet, Id_IndiceProjet, Fils, NbDoc, False) = False Then
                    DeletSheet SheetFiltreCible
                End If
                Else
                    If CrerSoutotal(SheetFiltreCible, Id_Onglet, Id_IndiceProjet, Fils, NbDoc, False, False) = False Then
                    DeletSheet SheetFiltreCible
                End If
                End If
            End If
            Next
    End If
End If
  While RsFiltre.EOF = False
  IsFilterElab = True
    Set SheetFiltre = IsertSheet(MyExel, "MyFilterElaborré")
    DeleteRow SheetFiltre, True
    ExcelCreatTitre SheetFiltre.Range("a1").CurrentRegion, Rs
    Set MyRange = SheetFiltre.Range("A1").CurrentRegion
 For I = 1 To MyRange.Columns.Count
 
 



         Sql = "SELECT T_Etats_Select_Filtre.FiltreName, T_Etats_Select_Filtre.Colonne, T_Etats_Select_Filtre.Valeur, T_Etats_Select_Filtre.Ligne "
        Sql = Sql & "FROM T_Etats_Onglet LEFT JOIN T_Etats_Select_Filtre ON T_Etats_Onglet.Id = T_Etats_Select_Filtre.Id_Onglet "
        Sql = Sql & "WHERE T_Etats_Select_Filtre.FiltreName='" & MyReplace(RsFiltre!FiltreName) & "' "
        Sql = Sql & "AND T_Etats_Select_Filtre.Colonne='" & MyReplace(MyRange(1, I)) & "' "
        Sql = Sql & "AND T_Etats_Select_Filtre.Id_Onglet=" & Id_Onglet & " "
        Sql = Sql & "ORDER BY T_Etats_Select_Filtre.Id;"

        
        

 Set RsFiltreELab = Con.OpenRecordSet(Sql)
    While RsFiltreELab.EOF = False
        
        
            
                MyRange(RsFiltreELab!Ligne, I) = "" & RsFiltreELab!Valeur
            SaveNbL = RsFiltreELab!Ligne
           
       
        RsFiltreELab.MoveNext
    Wend
Next

Set RsFiltreELab = Con.CloseRecordSet(RsFiltreELab)
For I = 2 To SaveNbL
    If SheetFiltre.Range("a1").CurrentRegion.Rows.Count = 1 Then
        SheetFiltre.Rows(CStr(2) & ":" & CStr(2)).Delete Shift:=xlUp
    
    End If
Next
Sql = "SELECT T_Etats_Select_Champs.ChampAs, T_indiceProjet.Equipement, T_Etats_Select_Champs.Id_Onglet "
        Sql = Sql & "FROM T_indiceProjet, T_Etats_Onglet INNER JOIN T_Etats_Select_Champs ON T_Etats_Onglet.Id = T_Etats_Select_Champs.Id_Onglet "
        Sql = Sql & "WHERE (T_Etats_Select_Champs.ChamsName='OPTION' OR T_Etats_Select_Champs.ChamsName='CRITERES') "
        If Fils = 0 Then
         Sql = Sql & "AND T_indiceProjet.Id=" & Id_IndiceProjet & " "
         Else
            Sql = Sql & "AND T_indiceProjet.Id=" & Fils & " "
        End If
        Sql = Sql & "AND T_Etats_Select_Champs.Id_Onglet=" & Id_Onglet & " AND T_Etats_Onglet.FiltreEquipement=True "
        Sql = Sql & "GROUP BY T_Etats_Select_Champs.ChamsName, T_Etats_Select_Champs.ChampAs, T_indiceProjet.Equipement, T_indiceProjet.Id, "
        Sql = Sql & "T_Etats_Select_Champs.Id_Onglet;"
        

Set RsFiltreELab = Con.OpenRecordSet(Sql)
If RsFiltreELab.EOF = False Then

MyOption = FunCodeCritaire(Id_IndiceProjet, "" & RsFiltreELab!Equipement)
    NbMyOption = 0
    
Set RangeCopy = SheetFiltre.Range("a1").CurrentRegion
    For I = 0 To UBound(MyOption)
    PoseFiterMyOption = SheetFiltre.Range("a1").CurrentRegion.Rows.Count
    If Trim("" & MyOption(I)) <> "" Then
        If NbMyOption > 0 Then
            RangeCopy.Copy
            Debut = PoseFiterMyOption + 1
            SheetFiltre.Select
            SheetFiltre.Cells(Debut, 1).Select
            SheetFiltre.Paste
            SheetFiltre.Rows(CStr(Debut) & ":" & CStr(Debut)).Delete Shift:=xlUp
            
        Else
            Debut = 2
        End If
     For I2 = Debut To SheetFiltre.Range("a1").CurrentRegion.Rows.Count
        For I3 = 1 To SheetFiltre.Range("a1").CurrentRegion.Columns.Count
            If SheetFiltre.Cells(1, I3) = RsFiltreELab!ChampAs Then
                SheetFiltre.Cells(I2, I3) = "*;" & UCase(Replace("" & MyOption(I), "§Null§", "") & ";*")
              
               
                Exit For
             End If
            
        Next
     Next
        NbMyOption = NbMyOption + 1
    End If
Next
End If
If Trim(LstOnglet) = "" Then
     Set SheetFiltreCible = IsertSheet(MyExel, MySeet.Name & "_FLT_" & RsFiltre!FiltreName, True)
    FiltreActif MySeet.Range("a1").CurrentRegion, SheetFiltre.Range("A1").CurrentRegion, SheetFiltreCible.Range("a1").CurrentRegion
     OfseteMiseEnPage = SurEquipement(PrefixName, MyWorkbook, SheetFiltreCible, Id_IndiceProjet, Id_Onglet, Fils)
    Set MyRange = SheetFiltreCible.Range("A1").CurrentRegion
       
                SheetFiltreCible.Cells.Replace ";", ""
                    RsChampMasquer.Requery
                    KillColImPaire = False
                While RsChampMasquer.EOF = False
                    SheetFiltreCible.Range(Replace(SheetFiltreCible.Cells(1, MyClonne("" & RsChampMasquer!ChamsName)).Address, "$", "")).Interior.ColorIndex = 6
                     RsChampMasquer.MoveNext
                Wend
                If DapresDoc = False Then
                    For I5 = SheetFiltreCible.Range("A1").CurrentRegion.Columns.Count To 1 Step -1
                        If SheetFiltreCible.Cells(1, I5).Interior.ColorIndex <> 6 Then
                            SuprmerCells SheetFiltreCible.Range(Replace(SheetFiltreCible.Cells(1, I5).Address & ":" & SheetFiltreCible.Cells(SheetFiltreCible.Range("A1").CurrentRegion.Rows.Count, I5).Address, "$", "")), "g"
                            If KillColImPaire = True Then
                                SuprmerCells SheetFiltreCible.Range(Replace(SheetFiltreCible.Cells(SheetFiltreCible.Range("A1").CurrentRegion.Rows.Count + 1, 1).Address & ":" & SheetFiltreCible.Cells(SheetFiltreCible.Range("A1").CurrentRegion.Rows.Count + 1000, 1).Address, "$", "")), "g"
                                KillColImPaire = False
                            Else
                                KillColImPaire = True
                            End If
                         End If
                     Next
               End If
    If CrerSoutotal(SheetFiltreCible, Id_Onglet, Id_IndiceProjet, Fils, NbDoc) = False Then
        DeletSheet SheetFiltreCible
    End If
Else
    SplitLstOnglet = Split(LstOnglet & ";", ";")
    For I = 2 To MySeet.Range("a1").CurrentRegion.Rows.Count
        MySeet.Cells(I, MyClonne(LstOngletName)) = ";" & MySeet.Cells(I, MyClonne(LstOngletName)) & ";"
        
    Next
   
    For I4 = 0 To UBound(SplitLstOnglet)
        If Trim(SplitLstOnglet(I4)) <> "" Then
            For I = 2 To SheetFiltre.Range("a1").CurrentRegion.Rows.Count
                SheetFiltre.Cells(I, MyClonne(LstOngletName)) = ";" & SplitLstOnglet(I4) & ";"
            Next
            Set SheetFiltreCible = IsertSheet(MyExel, MySeet.Name & "_FLT_" & RsFiltre!FiltreName & "_" & SplitLstOnglet(I4), True)
            FiltreActif MySeet.Range("a1").CurrentRegion, SheetFiltre.Range("A1").CurrentRegion, SheetFiltreCible.Range("a1").CurrentRegion
            OfseteMiseEnPage = SurEquipement(PrefixName, MyWorkbook, SheetFiltreCible, Id_IndiceProjet, Id_Onglet, Fils)
             For I = 2 To SheetFiltreCible.Range("a1").CurrentRegion.Rows.Count
                    SheetFiltreCible.Cells(I, MyClonne(LstOngletName)) = Replace(SheetFiltreCible.Cells(I, MyClonne(LstOngletName)), ";", "")
                Next
                
                    RsChampMasquer.Requery
                    KillColImPaire = False
                While RsChampMasquer.EOF = False
                    SheetFiltreCible.Range(Replace(SheetFiltreCible.Cells(1, MyClonne("" & RsChampMasquer!ChamsName)).Address, "$", "")).Interior.ColorIndex = 6
                     RsChampMasquer.MoveNext
                Wend
                If DapresDoc = False Then
                    For I5 = SheetFiltreCible.Range("A1").CurrentRegion.Columns.Count To 1 Step -1
                        If SheetFiltreCible.Cells(1, I5).Interior.ColorIndex <> 6 Then
                            SuprmerCells SheetFiltreCible.Range(Replace(SheetFiltreCible.Cells(1, I5).Address & ":" & SheetFiltreCible.Cells(SheetFiltreCible.Range("A1").CurrentRegion.Rows.Count, I5).Address, "$", "")), "g"
                            If KillColImPaire = True Then
                                SuprmerCells SheetFiltreCible.Range(Replace(SheetFiltreCible.Cells(SheetFiltreCible.Range("A1").CurrentRegion.Rows.Count + 1, 1).Address & ":" & SheetFiltreCible.Cells(SheetFiltreCible.Range("A1").CurrentRegion.Rows.Count + 1000, 1).Address, "$", "")), "g"
                                KillColImPaire = False
                            Else
                                KillColImPaire = True
                            End If
                         End If
                     Next
               End If
            Set MyRange = SheetFiltreCible.Range("A1").CurrentRegion
    
    
    
            For I = 2 To SheetFiltreCible.Range("a1").CurrentRegion.Rows.Count
            SheetFiltreCible.Cells(I, MyClonne(LstOngletName)) = Replace(SheetFiltreCible.Cells(I, MyClonne(LstOngletName)), ";", "")
            Next
           
 
       
         

If CrerSoutotal(SheetFiltreCible, Id_Onglet, Id_IndiceProjet, Fils, NbDoc, False) = False Then
    DeletSheet SheetFiltreCible
End If

 
        End If
    Next
     
End If


   
 
   

RsFiltre.MoveNext

Wend
If IsFilterElab = True Then
    MySeet.Delete
Else
    CrerSoutotal MySeet, Id_Onglet, Id_IndiceProjet, Fils, NbDoc
End If
Set CreatEtaExels = MySeet
End Function

Function CreerXls2DapresDoc(MyExel As EXCEL.Application, MyWorkbook As Workbook, MyWorkbookSource As Workbook, Rs As Recordset, Onglet As String, Trie As String, _
                        Id_Onglet As Long, Id_IndiceProjet As Long, Fils As Long, Id_Macro As Long, _
                        NbDoc As Long, Optional DapresDoc As Boolean) As Worksheet
Dim MySeet As EXCEL.Worksheet
Dim MyRange As EXCEL.Range
Dim SheetFiltre As Worksheet
Dim SheetFiltreCible As Worksheet
Dim SaveNbL As Long
Dim RsSelectChamp As Recordset
Dim RsCriteres As Recordset
Dim MyTrie
Dim MyTrieS
Dim Sql As String
Dim RsFiltre As Recordset
Dim RsFiltreELab As Recordset
Dim I As Long
Dim TrieExcel(3, 1) As String
Dim Col As Long
Dim NbTrie As Integer
Dim IsFilterElab As Boolean
Dim MyOption
Dim MyOption2
Dim NbMyOption As Long
Dim RangeCopy As Range
Dim PoseFiterMyOption As Long
Dim Debut As Long
Dim I2, I3, I4, I5 As Long
Dim L, C As Long
Dim OpTionValue
Dim RsFiterSurVue As Recordset
Dim ClsDog As GenerateurDoc
Dim SplitLstOnglet
Dim KillColImPaire As Boolean
Dim PasFiltre As Boolean
Dim PrefixName As String
Dim ChoixTraitement As Long
Dim PrefixNameSave As String
Dim Index As Long
PrefixName = "" ' "_FLT_S_Equ_"
StarOnglet = MyWorkbook.Sheets.Count
OfsetOnglet = -1
 Sql = "SELECT T_Etats_Select_Filtre.FiltreName "
Sql = Sql & "From T_Etats_Select_Filtre "
Sql = Sql & "Where T_Etats_Select_Filtre.Id_Onglet =  " & Id_Onglet & " "
Sql = Sql & "GROUP BY T_Etats_Select_Filtre.FiltreName "
Sql = Sql & "ORDER BY T_Etats_Select_Filtre.FiltreName; "
ChoixTraitement = 0
Set RsFiltre = Con.OpenRecordSet(Sql)
If RsFiltre.EOF = False Then ChoixTraitement = 1


    I = 0

'Select Case ChoixTraitement
'    Case 0
'        SurEquipement PrefixName, MyWorkbook, MySeet, Id_IndiceProjet, Id_Onglet, Fils
'         ChampMasquer RsChampMasquer, MySeet, DapresDoc
'         CrerSoutotal MySeet, Id_Onglet, Id_IndiceProjet, Fils, NbDoc
'    Case 1
'
'    Case 2
'            PrefixName = "_Ongt_"
'            SurOnglet PrefixName, MyWorkbook, RsChampMasquer, MySeet, Id_Onglet, Id_IndiceProjet, Fils, NbDoc, DapresDoc
'            MySeet.Delete
'    Case 3
'             While RsFiltre.EOF = False
'                PrefixName = "" & RsFiltre!FiltreName & "_"
'                If PrefixName <> PrefixNameSave And PrefixName <> "" Then
'                    SurOnglet PrefixName, MyWorkbook, RsChampMasquer, MySeet, Id_Onglet, Id_IndiceProjet, Fils, NbDoc, DapresDoc
'                End If
'                PrefixNameSave = PrefixName
'                 RsFiltre.MoveNext
'             Wend
'End Select
If Trim(LstOnglet) <> "" Then ChoixTraitement = ChoixTraitement + 2
 BoucleOnglets Id_Onglet, MyWorkbookSource
 If MyPlageOnglets.Err = False Then
    For Index = MyPlageOnglets.Fin To MyPlageOnglets.Debut Step -1
        Set MySeet = RetounSheet(MyWorkbook, MyWorkbookSource.Sheets(Index), Id_Onglet, GeneEtatMacro & "_")
        Select Case ChoixTraitement
            Case 0
            
                OfseteMiseEnPage = SurEquipement(PrefixName, MyWorkbook, MySeet, Id_IndiceProjet, Id_Onglet, Fils)
                 ChampMasquer RsChampMasquer, MySeet, DapresDoc
                 CrerSoutotal MySeet, Id_Onglet, Id_IndiceProjet, Fils, NbDoc
                 StarOnglet = 1
                 OfsetOnglet = OfsetOnglet + 1
            Case 1
                    PrefixName = ""
                    SurFitreElab PrefixName, MyWorkbook, MySeet, Id_IndiceProjet, Id_Onglet, Fils, NbDoc, DapresDoc
            Case 2
                    PrefixName = "_Ongt_"
                    SurOnglet PrefixName, MyWorkbook, RsChampMasquer, MySeet, Id_Onglet, Id_IndiceProjet, Fils, NbDoc, DapresDoc
                    MySeet.Delete
            Case 3
                     While RsFiltre.EOF = False
                        PrefixName = "" & RsFiltre!FiltreName & "_"
                        If PrefixName <> PrefixNameSave And PrefixName <> "" Then
                            SurOnglet PrefixName, MyWorkbook, RsChampMasquer, MySeet, Id_Onglet, Id_IndiceProjet, Fils, NbDoc, DapresDoc
                        End If
                        PrefixNameSave = PrefixName
                         RsFiltre.MoveNext
                     Wend
        End Select
        Set SheetFiltre = IsertSheet(MyWorkbook, "MyFilterElaborré")
        
    Next
    SheetFiltre.Delete
    StopOnglet = StarOnglet + OfsetOnglet
        DeletSheetEtat MyWorkbook, ActionType, NbDoc, StarOnglet, StopOnglet
End If
End Function
Function SurEquipement(PrefixName2 As String, MyExel As Workbook, MySeet As Worksheet, Id_IndiceProjet As Long, Id_Onglet As Long, Fils As Long)
Dim Sql As String
Dim RsFiltreELab  As Recordset
Dim IsFilterElab As Boolean
Dim MyOption
Dim SheetFiltre As Worksheet
Dim RsMacro As Recordset
Dim RsFiltre As Recordset
Dim RsSelectChamp As Recordset
Dim Table As String
Dim RsExcel  As Recordset
Dim I2, I3, C As Long
Dim PrefixName As String
Dim MyRange As Range
Dim SheetFiltreCible As Worksheet
  Set SheetFiltre = IsertSheet(MyExel, "MyFilterElaborré")
   
Sql = "SELECT T_Etats_Select_Champs.ChampAs, T_indiceProjet.Equipement, T_Etats_Select_Champs.Id_Onglet "
Sql = Sql & "FROM T_indiceProjet, T_Etats_Onglet INNER JOIN T_Etats_Select_Champs ON T_Etats_Onglet.Id = T_Etats_Select_Champs.Id_Onglet "
Sql = Sql & "WHERE (T_Etats_Select_Champs.ChamsName='OPTION' OR T_Etats_Select_Champs.ChamsName='CRITERES') "

'sql = "SELECT T_Etats_Select_Champs.ChampAs, T_indiceProjet.Equipement, T_Etats_Select_Champs.Id_Onglet "
'sql = sql & "FROM T_indiceProjet, T_Etats_Onglet INNER JOIN T_Etats_Select_Champs ON T_Etats_Onglet.Id = T_Etats_Select_Champs.Id_Onglet "
'sql = sql & "WHERE T_Etats_Select_Champs.Id_Onglet=" & Id_Onglet & "  "
'sql = sql & "AND T_indiceProjet.Id=864 "
'sql = sql & "AND (T_Etats_Select_Champs.ChamsName='OPTION'  "
'sql = sql & "or T_Etats_Select_Champs.ChamsName='CRITERES');"


If Fils = 0 Then
    Sql = Sql & "AND T_indiceProjet.Id=" & Id_IndiceProjet & " "
Else
    Sql = Sql & "AND T_indiceProjet.Id=" & Fils & " "
End If
    Sql = Sql & "AND T_Etats_Select_Champs.Id_Onglet=" & Id_Onglet & " AND T_Etats_Onglet.FiltreEquipement=True "
    Sql = Sql & "GROUP BY T_Etats_Select_Champs.ChamsName, T_Etats_Select_Champs.ChampAs, T_indiceProjet.Equipement, T_indiceProjet.Id, "
    Sql = Sql & "T_Etats_Select_Champs.Id_Onglet;"


Set RsFiltreELab = Con.OpenRecordSet(Sql)
If RsFiltreELab.EOF = False Then
    Sql = "SELECT T_Menu_Etat_Onglet.TableName "
    Sql = Sql & "FROM T_Etats_Onglet INNER JOIN T_Menu_Etat_Onglet ON T_Etats_Onglet.Onglet = T_Menu_Etat_Onglet.Menu "
    Sql = Sql & "WHERE T_Etats_Onglet.Id=" & Id_Onglet & ";"
    Set RsMacro = Con.OpenRecordSet(Sql)
    'While RsFiltre.EOF = False
    IsFilterElab = True
    Set SheetFiltre = IsertSheet(MyExel, "MyFilterElaborréSurEqui")
    DeleteRow SheetFiltre, True
    
    Sql = "SELECT T_Etats_Select_Champs.ChamsName, T_Etats_Select_Champs.ChampAs, T_Etats_Select_Champs.Trie,T_Etats_Select_Champs.Visible "
    Sql = Sql & "From T_Etats_Select_Champs "
    Sql = Sql & "Where T_Etats_Select_Champs.Id_Onglet =" & Id_Onglet & " "
    Sql = Sql & "ORDER BY T_Etats_Select_Champs.Id;"
    
    
    Set RsSelectChamp = Con.OpenRecordSet(Sql)
    If RsSelectChamp.EOF = False Then
        Table = "" & RsMacro!TableName
        Sql = "Select "
        While RsSelectChamp.EOF = False
        '        If RsSelectChamp!Trie <> 0 Then
        '            Trie = Trie & "§" & RsSelectChamp!ChamsName & "=" & RsSelectChamp!Trie
        '        End If
            Sql = Sql & "[" & Table & "].[" & RsSelectChamp!ChamsName & "]"
            If UCase(RsSelectChamp!ChamsName) <> UCase(RsSelectChamp!ChampAs) Then
                Sql = Sql & " AS [" & RsSelectChamp!ChampAs & "]"
            End If
            Sql = Sql & ","
            RsSelectChamp.MoveNext
        Wend
    End If
    Sql = Left(Sql, Len(Sql) - 1)
    Sql = Sql & " FROM " & Table & " Where " & Table & ".Id_IndiceProjet=" & Id_IndiceProjet & "; "
    
    Set RsExcel = Con.OpenRecordSet(Sql)
    ExcelCreatTitre SheetFiltre.Range("a1").CurrentRegion, RsExcel
    
    IsFilterElab = True
    MyOption = FunCodeCritaire(Id_IndiceProjet, "" & RsFiltreELab!Equipement)
    Set SheetFiltre = IsertSheet(MyExel, "MyFilterElaborréSurEqui")
    DeleteRow SheetFiltre, True
    ExcelCreatTitre SheetFiltre.Range("a1").CurrentRegion, RsExcel
    Sql = "SELECT T_Etats_Select_Champs.ChampAs, T_Etats_Select_Champs.CreatOnglet "
    Sql = Sql & "FROM T_Etats_Onglet INNER JOIN T_Etats_Select_Champs ON T_Etats_Onglet.Id = T_Etats_Select_Champs.Id_Onglet  "
    Sql = Sql & "Where T_Etats_Select_Champs.Id_Onglet = " & RsFiltreELab!Id_Onglet & "  "
    Sql = Sql & " And T_Etats_Select_Champs.CreatOnglet=True "
    Sql = Sql & "GROUP BY T_Etats_Select_Champs.ChampAs, T_Etats_Select_Champs.CreatOnglet, T_Etats_Select_Champs.ChamsName;"
    
    For I2 = 0 To UBound(MyOption)
        For I3 = 1 To SheetFiltre.Range("a1").CurrentRegion.Columns.Count
            If SheetFiltre.Cells(1, I3) = RsFiltreELab!ChampAs Then
                PrefixName = "_FLT_S_Equ"
                SheetFiltre.Cells(I2 + 2, I3) = "*;" & UCase(Replace("" & MyOption(I2), "§Null§", "") & ";*")
                Exit For
            End If
        Next
    Next
'    "§Null§"
SheetFiltre.Range("a1").CurrentRegion.Replace ";;", "§Null§"
    Set MyRange = MySeet.Range("a1").CurrentRegion
    For I2 = 1 To MyRange.Columns.Count
        If MyRange(1, I2) = "OPTION" Then Exit For
    Next
    For I3 = 2 To MyRange.Rows.Count
        If Trim("" & MyRange(I3, I2)) = "" Then MyRange(I3, I2) = ";;"
    Next
    
    OfsetOnglet = OfsetOnglet + 1
    Set SheetFiltreCible = IsertSheet(MyExel, MySeet.Name & PrefixName, True)
    FiltreActif MySeet.Range("a1").CurrentRegion, SheetFiltre.Range("A1").CurrentRegion, SheetFiltreCible.Range("a1").CurrentRegion
    
    MySeet.Delete
    SheetFiltre.Delete
Set MySeet = SheetFiltreCible
End If
If BoolGenEtatEpisure = True Then
    VueArriere SheetFiltreCible
 End If
End Function
Function CreerXls2(MyExel As EXCEL.Application, MyWorkbook As Workbook, Rs As Recordset, Onglet As String, Trie As String, _
                        Id_Onglet As Long, Id_IndiceProjet As Long, Fils As Long, Id_Macro As Long, _
                        NbDoc As Long, Optional DapresDoc As Boolean) As Worksheet
Dim MySeet As EXCEL.Worksheet

Dim MyRange As EXCEL.Range
Dim SheetFiltre As Worksheet
Dim SheetFiltreCible As Worksheet
Dim SaveNbL As Long
Dim RsSelectChamp As Recordset
Dim RsCriteres As Recordset
Dim MyTrie
Dim MyTrieS
Dim Sql As String
Dim RsFiltre As Recordset
Dim RsFiltreELab As Recordset
Dim I As Long
Dim TrieExcel(3, 1) As String
Dim Col As Long
Dim NbTrie As Integer
Dim IsFilterElab As Boolean
Dim MyOption
Dim MyOption2
Dim NbMyOption As Long
Dim RangeCopy As Range
Dim PoseFiterMyOption As Long
Dim Debut As Long
Dim I2, I3, I4, I5 As Long
Dim L, C As Long
Dim OpTionValue
Dim RsFiterSurVue As Recordset
Dim ClsDog As GenerateurDoc
Dim SplitLstOnglet
'Dim RsChampMasquer As Recordset
Dim KillColImPaire As Boolean
Dim PasFiltre As Boolean
Dim PrefixName As String
Dim ChoixTraitement As Long
Dim PrefixNameSave As String
PrefixName = "" ' "_FLT_S_Equ_"
Set MySeet = IsertSheet(MyWorkbook, Onglet)
Set MyRange = MySeet.Range("A1").CurrentRegion

ExcelCreatTitre MyRange, Rs


 Sql = "SELECT DISTINCT T_Etats_Onglet.Id AS Id_Select, T_Etats_Select_Champs.ChamsName, T_Etats_Select_Champs.Visible "
    Sql = Sql & "FROM T_Etats_Onglet INNER JOIN T_Etats_Select_Champs ON T_Etats_Onglet.Id = T_Etats_Select_Champs.Id_Onglet "
    Sql = Sql & "Where T_Etats_Onglet.Id = " & Id_Onglet & " "
    Sql = Sql & "And T_Etats_Select_Champs.Visible = false; "
    Set RsChampMasquer = Con.OpenRecordSet(Sql)
'
' If Rs.EOF = False Then
'
'    MySeet.Range("A2").CopyFromRecordset Rs
'End If
If Rs.EOF = False Then
MySeet.Range("A2").CopyFromRecordset Rs
      ReplaceBool MySeet, "A2:" & Replace(MySeet.Range(Replace(MySeet.Cells(MySeet.Range("A1").CurrentRegion.Rows.Count, MySeet.Range("A1").CurrentRegion.Columns.Count).Address, "$", "")).Address, "$", "")
End If
'While Rs.EOF = False
'    Set Myrange = MySeet.Range("A1").CurrentRegion
'    Set Myrange = MySeet.Range(Myrange(Myrange.Rows.Count + 1, 1).Address & ":" & Myrange(Myrange.Rows.Count + 1, Myrange.Columns.Count).Address)
'    ExcelCreatTitre Myrange, Rs, True, Formule:=True
'    Rs.MoveNext
' Wend

RenseigneColonn MySeet, Id_Onglet, 1
 MyTrie = Split(Trie & "§", "§")
 If UBound(MyTrie) > 1 Then
    For I = 1 To UBound(MyTrie)
        MyTrieS = Split(MyTrie(I) & "=", "=")
    If Val(MyTrieS(1)) <> 0 Then
    TrieExcel(I, 0) = MyTrieS(0): TrieExcel(I, 1) = MyTrieS(1)
    If I = 3 Then Exit For
    End If
   
 Next
 Set MyRange = MySeet.Range("A1").CurrentRegion
NbTrie = 0
 For I = 1 To 3
    For Col = 1 To MyRange.Columns.Count
       If MyRange(1, Col) = TrieExcel(I, 0) Then
       NbTrie = NbTrie + 1
           TrieExcel(I, 0) = MyRange(2, Col).Address
            TrieExcel(I, 0) = Replace(TrieExcel(I, 0), "$", "")
           Exit For
       End If
    Next
  Next
 End If
 Trier MySeet, NbTrie, Replace("A2:" & MyRange(MyRange.Rows.Count, MyRange.Columns.Count).Address, "$", ""), TrieExcel(1, 0), Val(TrieExcel(1, 1)), TrieExcel(2, 0), Val(TrieExcel(2, 1)), TrieExcel(3, 0), Val(TrieExcel(3, 1))
 Set MyRange = MySeet.Range("A1").CurrentRegion
 For I = 1 To MyRange.Columns.Count
    If UCase(MyRange(1, I)) = "OPTION" Or UCase(MyRange(1, I)) = UCase("Criteres") Then
        For I2 = 2 To MyRange.Rows.Count
            MyRange(I2, I) = ";" & MyRange(I2, I) & ";"
            
        Next
        Exit For
    End If
 Next
 MySeet.Range("a1").CurrentRegion.Replace ";;", "§Null§"
 Set CollecFiltreName = Nothing
 Set CollecFiltreName = New Collection
 Sql = "SELECT T_Etats_Select_Filtre.FiltreName "
Sql = Sql & "From T_Etats_Select_Filtre "
Sql = Sql & "Where T_Etats_Select_Filtre.Id_Onglet =  " & Id_Onglet & " "
Sql = Sql & "GROUP BY T_Etats_Select_Filtre.FiltreName "
Sql = Sql & "ORDER BY T_Etats_Select_Filtre.FiltreName; "

Set RsFiltre = Con.OpenRecordSet(Sql)
If RsFiltre.EOF = False Then ChoixTraitement = 1


    I = 0
If Trim(LstOnglet) <> "" Then ChoixTraitement = ChoixTraitement + 2
Select Case ChoixTraitement
    Case 0
        OfseteMiseEnPage = SurEquipement(PrefixName, MyWorkbook, MySeet, Id_IndiceProjet, Id_Onglet, Fils)
         ChampMasquer RsChampMasquer, MySeet, DapresDoc
         CrerSoutotal MySeet, Id_Onglet, Id_IndiceProjet, Fils, NbDoc, OfseteMiseEnPage:=OfseteMiseEnPage, EnteteOnglet:=Trim("" & Onglet)
        
    Case 1
            PrefixName = ""
            SurFitreElab PrefixName, MyWorkbook, MySeet, Id_IndiceProjet, Id_Onglet, Fils, NbDoc, DapresDoc
             MySeet.Delete
    Case 2
            PrefixName = "_Ongt_"
            SurOnglet PrefixName, MyWorkbook, RsChampMasquer, MySeet, Id_Onglet, Id_IndiceProjet, Fils, NbDoc, DapresDoc
            MySeet.Delete
    Case 3
             While RsFiltre.EOF = False
                PrefixName = "" & RsFiltre!FiltreName & "_"
                If PrefixName <> PrefixNameSave And PrefixName <> "" Then
                    SurOnglet PrefixName, MyWorkbook, RsChampMasquer, MySeet, Id_Onglet, Id_IndiceProjet, Fils, NbDoc, DapresDoc
                End If
                PrefixNameSave = PrefixName
                 RsFiltre.MoveNext
             Wend
End Select
Set SheetFiltre = IsertSheet(MyWorkbook, "MyFilterElaborré")
SheetFiltre.Delete

End Function
Sub MarquerChampSup(Champ As String, MySeet As Worksheet)
On Error GoTo Fin
MySeet.Range(Replace(MySeet.Cells(1, MyClonne("" & Champ)).Address, "$", "")).Select
    MySeet.Range(Replace(MySeet.Cells(1, MyClonne("" & Champ)).Address, "$", "")).Interior.ColorIndex = 6
Fin:
Err.Clear
End Sub
Sub ChampMasquer(RsChampMasquer As Recordset, MySeet As Worksheet, DapresDoc As Boolean)
Dim I As Long
Dim MyRange As Range
Dim KillColImPaire As Boolean
MySeet.Cells.Replace ";", ""
If BoolGenEtatEpisure = True Then
    OfseteMiseEnPage = VueArriere(MySeet)
   End If
RsChampMasquer.Requery

While RsChampMasquer.EOF = False
'                SheetFiltreCible.Range("A1").Interior.ColorIndex
    Debug.Print MySeet.Parent.FullName
    MySeet.Application.Workbooks(MySeet.Parent.Name).Activate
    MySeet.Activate
    MarquerChampSup "" & RsChampMasquer!ChamsName, MySeet
    RsChampMasquer.MoveNext
Wend
If DapresDoc = False Then
RenameTitre MySeet
Set MyRange = MySeet.Cells(1, MyClonne(1)).CurrentRegion
Set MyRange = MySeet.Range("A1:" & Replace(MySeet.Cells(MyRange.Rows.Count, MyRange.Columns.Count).Address, "$", ""))
    For I = MyRange.Columns.Count To 1 Step -1
      
        If MySeet.Cells(1, I).Interior.ColorIndex <> 6 Then
            SuprmerCells MySeet.Range(Replace(MySeet.Cells(1, I).Address & ":" & MySeet.Cells(MySeet.Range("A1").CurrentRegion.Rows.Count, I).Address, "$", "")), "g"
            If MySeet.Range("A1").CurrentRegion.Columns.Count > 2 Then
                If KillColImPaire = True Then
                    SuprmerCells MySeet.Range(Replace(MySeet.Cells(MySeet.Range("A1").CurrentRegion.Rows.Count + 1, 1).Address & ":" & MySeet.Cells(MySeet.Range("A1").CurrentRegion.Rows.Count + 1000, 1).Address, "$", "")), "g"
                    KillColImPaire = False
                Else
                    KillColImPaire = True
                End If
                End If
         End If
         Next

End If
                
End Sub

Sub SurOnglet(PrefixName As String, MyExel As EXCEL.Workbook, RsChampMasquer As Recordset, MySheet As Worksheet, Id_Onglet As Long, Id_IndiceProjet As Long, Fils As Long, NbDoc As Long, DapresDoc As Boolean, Optional BoolFill As Boolean = True, Optional Executer As Boolean = True)
Dim SplitLstOnglet
Dim I, I4 As Long
Dim C2 As Long
Dim SheetFiltre As Worksheet
Dim SheetFiltreCible As Worksheet
Dim NotNulls As Boolean
'Dim PrefixName As String
Dim C As Long

SplitLstOnglet = Split(LstOnglet & ";", ";")
    For I = 2 To MySheet.Range("a1").CurrentRegion.Rows.Count
        MySheet.Cells(I, MyClonne(LstOngletName)) = ";" & MySheet.Cells(I, MyClonne(LstOngletName)) & ";"
    Next
            For I4 = 0 To UBound(SplitLstOnglet)
            If Trim(SplitLstOnglet(I4)) <> "" Then
            Set SheetFiltre = IsertSheet(MyExel, "MyFilterElaborré")
            If SheetFiltre.Range("A1").CurrentRegion.Rows.Count = 1 And SheetFiltre.Range("A1").CurrentRegion.Columns.Count = 1 Then
                If Trim("" & SheetFiltre.Range("A1")) = "" Then
                    
                    For C = 1 To MySheet.Range("a1").CurrentRegion.Columns.Count
                        SheetFiltre.Cells(1, C) = MySheet.Cells(1, C)
                        SheetFiltre.Cells(2, MyClonne(LstOngletName)) = "a"
                    Next
                End If
            End If

                For I = 2 To SheetFiltre.Range("a1").CurrentRegion.Rows.Count
                    SheetFiltre.Cells(I, MyClonne(LstOngletName)) = ";" & SplitLstOnglet(I4) & ";"
                Next
                  
                        Set SheetFiltreCible = IsertSheet(MyExel, MySheet.Name & PrefixName & SplitLstOnglet(I4), True)
                 
             
                FiltreActif MySheet.Range("a1").CurrentRegion, SheetFiltre.Range("A1").CurrentRegion, SheetFiltreCible.Range("a1").CurrentRegion
                
                NotNulls = True
                If SheetFiltreCible.Range("A1").CurrentRegion.Rows.Count = 2 Then
'                    MsgBox "D_Ongt_120.KB"
                    NotNulls = False
                    For C2 = 1 To MyClonne.Count
                    If Trim("" & SheetFiltreCible.Cells(2, MyClonne(C2))) <> "" Then
                        SheetFiltreCible.Cells(2, MyClonne(C2)).Select
                        If InStr(1, UCase(SheetFiltreCible.Name), UCase(Replace("" & SheetFiltreCible.Cells(2, MyClonne(C2)), ";", ""))) = 0 And UCase(SheetFiltreCible.Cells(1, MyClonne(C2))) <> "ACTIVER" Then
                            NotNulls = True
                        End If
                        End If
                       
                    Next
                End If
                If NotNulls = False Then
                    SheetFiltreCible.Rows(2).Delete

                End If
            OfseteMiseEnPage = SurEquipement(PrefixName, MyExel, SheetFiltreCible, Id_IndiceProjet, Id_Onglet, Fils)
                SheetFiltreCible.Range("A1").CurrentRegion.Replace ";", ""
                 SheetFiltreCible.Range("A1").CurrentRegion.Replace "§Null§", ""
                
                    ChampMasquer RsChampMasquer, SheetFiltreCible, DapresDoc
                CrerSoutotal SheetFiltreCible, Id_Onglet, Id_IndiceProjet, Fils, NbDoc, DapresDoc, OfseteMiseEnPage:=OfseteMiseEnPage, EnteteOnglet:="" & SplitLstOnglet(I4)
            End If
            Next
    
End Sub
Function SurFitreElab(PrefixName2 As String, MyExel As Workbook, MySeet As Worksheet, Id_IndiceProjet As Long, Id_Onglet As Long, Fils As Long, NbDoc As Long, DapresDoc As Boolean)
Dim Sql As String
Dim RsFiltreELab  As Recordset
Dim IsFilterElab As Boolean
Dim MyOption
Dim SheetFiltre As Worksheet
Dim RsMacro As Recordset
Dim RsFiltre As Recordset
Dim RsSelectChamp As Recordset
Dim Table As String
Dim RsExcel  As Recordset
Dim I2, I3, C As Long
Dim PrefixName As String
Dim MyRange As Range
Dim SheetFiltreCible As Worksheet
Dim SaveFiltreName As String
Dim L As Long
'  Set SheetFiltre = IsertSheet(MyExel, "MyFilterElaborré")
   
Sql = "SELECT T_Etats_Onglet.Id, T_Etats_Select_Filtre.FiltreName, T_Etats_Select_Filtre.Colonne,  "
    Sql = Sql & "T_Etats_Select_Filtre.Valeur "
    Sql = Sql & "FROM T_Etats_Onglet INNER JOIN T_Etats_Select_Filtre ON T_Etats_Onglet.Id =  "
    Sql = Sql & "T_Etats_Select_Filtre.Id_Onglet "
    Sql = Sql & "WHERE T_Etats_Onglet.Id=" & Id_Onglet & ";"
    
    
    



Set RsFiltreELab = Con.OpenRecordSet(Sql)
If RsFiltreELab.EOF = False Then
    Sql = "SELECT T_Menu_Etat_Onglet.TableName "
    Sql = Sql & "FROM T_Etats_Onglet INNER JOIN T_Menu_Etat_Onglet ON T_Etats_Onglet.Onglet = T_Menu_Etat_Onglet.Menu "
    Sql = Sql & "WHERE T_Etats_Onglet.Id=" & Id_Onglet & ";"
    Set RsMacro = Con.OpenRecordSet(Sql)
    'While RsFiltre.EOF = False
    IsFilterElab = True
    

    Sql = "SELECT T_Etats_Select_Champs.ChamsName, T_Etats_Select_Champs.ChampAs, T_Etats_Select_Champs.Trie,T_Etats_Select_Champs.Visible "
    Sql = Sql & "From T_Etats_Select_Champs "
    Sql = Sql & "Where T_Etats_Select_Champs.Id_Onglet =" & Id_Onglet & " "
    Sql = Sql & "ORDER BY T_Etats_Select_Champs.Id;"


    Set RsSelectChamp = Con.OpenRecordSet(Sql)
    If RsSelectChamp.EOF = False Then
        Table = "" & RsMacro!TableName
'        sql = "Select "
'        While RsSelectChamp.EOF = False
'        '        If RsSelectChamp!Trie <> 0 Then
'        '            Trie = Trie & "§" & RsSelectChamp!ChamsName & "=" & RsSelectChamp!Trie
'        '        End If
'            sql = sql & "[" & Table & "].[" & RsSelectChamp!ChamsName & "]"
'            If UCase(RsSelectChamp!ChamsName) <> UCase(RsSelectChamp!ChampAs) Then
'                sql = sql & " AS [" & RsSelectChamp!ChampAs & "]"
'            End If
'            sql = sql & ","
'            RsSelectChamp.MoveNext
'        Wend
    End If
    
    Sql = "select " & Table & ".*  "
    Sql = Sql & " FROM " & Table & " Where " & Table & ".Id_IndiceProjet=" & Id_IndiceProjet & "; "

    Set RsExcel = Con.OpenRecordSet(Sql)
    While RsFiltreELab.EOF = False
        IsFilterElab = True
        If SaveFiltreName <> "" And SaveFiltreName <> "" & RsFiltreELab!FiltreName Then
             Set SheetFiltreCible = IsertSheet(MyExel, "S_Filtre_" & MySeet.Name & "_" & SaveFiltreName, True)
             FiltreActif MySeet.Range("a1").CurrentRegion, SheetFiltre.Range("A1").CurrentRegion, SheetFiltreCible.Range("a1").CurrentRegion
                If BoolGenEtatEpisure = True Then
                    VueArriere SheetFiltreCible
                End If
                ChampMasquer SheetFiltreCible, SheetFiltreCible, DapresDoc
                CrerSoutotal MySeet, Id_Onglet, Id_IndiceProjet, Fils, NbDoc

            SaveFiltreName = ""
        Else
        If SaveFiltreName = "" Then
        Set SheetFiltre = IsertSheet(MyExel, "MyFilterElaborré")
            DeleteRow SheetFiltre, True
            For C = 1 To MySeet.Range("a1").CurrentRegion.Columns.Count
                SheetFiltre.Cells(1, C) = MySeet.Cells(1, C)
            Next
'            ExcelCreatTitre SheetFiltre.Range("a1").CurrentRegion, RsExcel
             SaveFiltreName = "" & RsFiltreELab!FiltreName
        End If
         
        End If
        If SaveFiltreName <> "" Then
            SheetFiltre.Cells(SheetFiltre.Range("A1").CurrentRegion.Rows.Count + 1, MyClonne("" & RsFiltreELab!Colonne)) = "'" & "" & RsFiltreELab!Valeur
        
         RsFiltreELab.MoveNext
         End If
     
    Wend
End If
If IsFilterElab = True Then
    Set SheetFiltreCible = IsertSheet(MyExel, "S_Filtre_" & MySeet.Name & "_" & SaveFiltreName, True)
    FiltreActif MySeet.Range("a1").CurrentRegion, SheetFiltre.Range("A1").CurrentRegion, SheetFiltreCible.Range("a1").CurrentRegion
    If BoolGenEtatEpisure = True Then
        VueArriere SheetFiltreCible
    End If
    ChampMasquer RsChampMasquer, SheetFiltreCible, DapresDoc
    CrerSoutotal SheetFiltreCible, Id_Onglet, Id_IndiceProjet, Fils, NbDoc
End If
Set SheetFiltre = IsertSheet(MyExel, "MyFilterElaborré")
SheetFiltre.Delete
'    IsFilterElab = True
'    MyOption = FunCodeCritaire(Id_IndiceProjet, "" & RsFiltreELab!Equipement)
'    Set SheetFiltre = IsertSheet(MyExel, "MyFilterElaborréSurEqui")
'    DeleteRow SheetFiltre, True
'    ExcelCreatTitre SheetFiltre.Range("a1").CurrentRegion, RsExcel
'    sql = "SELECT T_Etats_Select_Champs.ChampAs, T_Etats_Select_Champs.CreatOnglet "
'    sql = sql & "FROM T_Etats_Onglet INNER JOIN T_Etats_Select_Champs ON T_Etats_Onglet.Id = T_Etats_Select_Champs.Id_Onglet  "
'    sql = sql & "Where T_Etats_Select_Champs.Id_Onglet = " & RsFiltreELab!Id_Onglet & "  "
'    sql = sql & " And T_Etats_Select_Champs.CreatOnglet=True "
'    sql = sql & "GROUP BY T_Etats_Select_Champs.ChampAs, T_Etats_Select_Champs.CreatOnglet, T_Etats_Select_Champs.ChamsName;"
'
'    For I2 = 0 To UBound(MyOption)
'        For I3 = 1 To SheetFiltre.Range("a1").CurrentRegion.Columns.Count
'            If SheetFiltre.Cells(1, I3) = RsFiltreELab!ChampAs Then
'                PrefixName = "_FLT_S_Equ"
'                SheetFiltre.Cells(I2 + 2, I3) = "*;" & UCase(Replace("" & MyOption(I2), "§Null§", "") & ";*")
'                Exit For
'            End If
'        Next
'    Next
'    Set MyRange = MySeet.Range("a1").CurrentRegion
'    For I2 = 1 To MyRange.Columns.Count
'        If MyRange(1, I2) = "OPTION" Then Exit For
'    Next
'    For I3 = 2 To MyRange.Rows.Count
'        If Trim("" & MyRange(I3, I2)) = "" Then MyRange(I3, I2) = ";;"
'    Next
'
'
'
'    Set SheetFiltreCible = IsertSheet(MyExel, MySeet.Name & PrefixName, True)
'    FiltreActif MySeet.Range("a1").CurrentRegion, SheetFiltre.Range("A1").CurrentRegion, SheetFiltreCible.Range("a1").CurrentRegion
'
'    MySeet.Delete
'    SheetFiltre.Delete
'Set MySeet = SheetFiltreCible
'End If

End Function
'Sub SurFitreElab(PrefixName As String, MyExel As EXCEL.Workbook, RsChampMasquer As Recordset, _
'                    RsFiltre As Recordset, MySheet As Worksheet, Id_Onglet As Long, Id_IndiceProjet As Long, Fils As Long, NbDoc As Long, _
'                    DapresDoc As Boolean, Trie As String, Optional BoolFill As Boolean = True, Optional Executer As Boolean = True)
'Dim IsFilterElab As Boolean
'Dim SheetFiltre As Worksheet
'Dim SheetFiltreCible As Worksheet
'Dim C, I, I5 As Long
'Dim MyRange As Range
'Dim KillColImPaire As Boolean
'Dim NameFiltre As String
'Dim sql As String
'Dim RsSelectChamp As Recordset
'Dim RsMacro As Recordset
'Dim Table As String
'Dim RsExcel As Recordset
'Dim RsFiltreELab As Recordset
'Dim SaveNbL
'sql = "SELECT T_Menu_Etat_Onglet.TableName "
'    sql = sql & "FROM T_Etats_Onglet INNER JOIN T_Menu_Etat_Onglet ON T_Etats_Onglet.Onglet = T_Menu_Etat_Onglet.Menu "
'    sql = sql & "WHERE T_Etats_Onglet.Id=" & Id_Onglet & ";"
'Set RsMacro = Con.OpenRecordSet(sql)
'While RsFiltre.EOF = False
'  IsFilterElab = True
'    Set SheetFiltre = IsertSheet(MyExel, "MyFilterElaborré")
'    DeleteRow SheetFiltre, True
'    sql = "SELECT T_Etats_Select_Champs.ChamsName, T_Etats_Select_Champs.ChampAs, T_Etats_Select_Champs.Trie,T_Etats_Select_Champs.Visible "
'    sql = sql & "From T_Etats_Select_Champs "
'    sql = sql & "Where T_Etats_Select_Champs.Id_Onglet =" & Id_Onglet & " "
'    sql = sql & "ORDER BY T_Etats_Select_Champs.Id;"
'
'
'    Set RsSelectChamp = Con.OpenRecordSet(sql)
'    If RsSelectChamp.EOF = False Then
'            Table = "" & RsMacro!TableName
'            sql = "Select "
'            Trie = ""
'        While RsSelectChamp.EOF = False
'        If RsSelectChamp!Trie <> 0 Then
'            Trie = Trie & "§" & RsSelectChamp!ChamsName & "=" & RsSelectChamp!Trie
'        End If
'            sql = sql & "[" & Table & "].[" & RsSelectChamp!ChamsName & "]"
'            If UCase(RsSelectChamp!ChamsName) <> UCase(RsSelectChamp!ChampAs) Then
'            sql = sql & " AS [" & RsSelectChamp!ChampAs & "]"
'            End If
'            sql = sql & ","
'            RsSelectChamp.MoveNext
'        Wend
'        sql = Left(sql, Len(sql) - 1)
'        sql = sql & " FROM " & Table & " Where " & Table & ".Id_IndiceProjet=" & Id_IndiceProjet & "; "
'
'        Set RsExcel = Con.OpenRecordSet(sql)
'         ExcelCreatTitre SheetFiltre.Range("a1").CurrentRegion, RsExcel
'    Set MyRange = SheetFiltre.Range("A1").CurrentRegion
''        CreatEtaExels MyWorkbook, RsExcel, "" & RsMacro!SaveOnglet, "" & Trie, RsMacro!Id_Select, Id_IndiceProjet, Fils, Id_Onglet, NbDoc
'
'        Set RsExcel = Con.CloseRecordSet(RsExcel)
'     End If
'
'    Set RsSelectChamp = Con.CloseRecordSet(RsSelectChamp)
''    For C = 1 To MySheet.Range("A1").CurrentRegion.Columns.Count
''    SheetFiltre.Cells(1, C) = MySheet.Cells(1, C)
''    Next
'    Set MyRange = SheetFiltre.Range("A1").CurrentRegion
'For I = 1 To MyRange.Columns.Count
'
'
'
'         sql = "SELECT T_Etats_Select_Filtre.FiltreName, T_Etats_Select_Filtre.Colonne, T_Etats_Select_Filtre.Valeur, T_Etats_Select_Filtre.Ligne "
'        sql = sql & "FROM T_Etats_Onglet LEFT JOIN T_Etats_Select_Filtre ON T_Etats_Onglet.Id = T_Etats_Select_Filtre.Id_Onglet "
'        sql = sql & "WHERE T_Etats_Select_Filtre.FiltreName='" & MyReplace(RsFiltre!FiltreName) & "' "
'        sql = sql & "AND T_Etats_Select_Filtre.Colonne='" & MyReplace(MyRange(1, I)) & "' "
'        sql = sql & "AND T_Etats_Select_Filtre.Id_Onglet=" & Id_Onglet & " "
'        sql = sql & "ORDER BY T_Etats_Select_Filtre.Id;"
'
'
'
'
' Set RsFiltreELab = Con.OpenRecordSet(sql)
'    While RsFiltreELab.EOF = False
'
'
'
'                MyRange(RsFiltreELab!Ligne, I) = "" & RsFiltreELab!Valeur
'            SaveNbL = RsFiltreELab!Ligne
'
'
'        RsFiltreELab.MoveNext
'    Wend
'Next
'
'Set RsFiltreELab = Con.CloseRecordSet(RsFiltreELab)
'For I = 2 To SaveNbL
'    If SheetFiltre.Range("a1").CurrentRegion.Rows.Count = 1 Then
'        SheetFiltre.Rows(CStr(2) & ":" & CStr(2)).Delete Shift:=xlUp
'
'    End If
'Next
'
'
'
'     Set SheetFiltreCible = IsertSheet(MyExel, MySheet.Name & "_FLT_" & RsFiltre!FiltreName, True)
'    FiltreActif MySheet.Range("a1").CurrentRegion, SheetFiltre.Range("A1").CurrentRegion, SheetFiltreCible.Range("a1").CurrentRegion
'    SurEquipement PrefixName, MyExel, SheetFiltreCible, Id_IndiceProjet, Id_Onglet, Fils
'    Set MyRange = SheetFiltreCible.Range("A1").CurrentRegion
'      ChampMasquer RsChampMasquer, SheetFiltreCible, DapresDoc
'                CrerSoutotal SheetFiltreCible, Id_Onglet, Id_IndiceProjet, Fils, NbDoc, DapresDoc
'''       FiltreActif MySheet.Range("a1").CurrentRegion, SheetFiltre.Range("A1").CurrentRegion, SheetFiltreCible.Range("a1").CurrentRegion
''            SurEquipement PrefixName, MyExel, SheetFiltreCible, Id_IndiceProjet, Id_Onglet, Fils
''                SheetFiltreCible.Range("A1").CurrentRegion.Replace ";", ""
''                SheetFiltreCible.Cells.Replace ";", ""
''                    RsChampMasquer.Requery
''                    KillColImPaire = False
''                While RsChampMasquer.EOF = False
'''                SheetFiltreCible.Range("A1").Interior.ColorIndex
''                    SheetFiltreCible.Range(Replace(SheetFiltreCible.Cells(1, MyClonne("" & RsChampMasquer!ChamsName)).Address, "$", "")).Interior.ColorIndex = 6
''                     RsChampMasquer.MoveNext
''                Wend
''                If DapresDoc = False Then
''
'''                    For I5 = SheetFiltreCible.Range("A1").CurrentRegion.Columns.Count To 1 Step -1
'''                        If SheetFiltreCible.Cells(1, I5).Interior.ColorIndex <> 6 Then
'''                            SuprmerCells SheetFiltreCible.Range(Replace(SheetFiltreCible.Cells(1, I5).Address & ":" & SheetFiltreCible.Cells(SheetFiltreCible.Range("A1").CurrentRegion.Rows.Count, I5).Address, "$", "")), "g"
'''                            If KillColImPaire = True Then
'''                                SuprmerCells SheetFiltreCible.Range(Replace(SheetFiltreCible.Cells(SheetFiltreCible.Range("A1").CurrentRegion.Rows.Count + 1, 1).Address & ":" & SheetFiltreCible.Cells(SheetFiltreCible.Range("A1").CurrentRegion.Rows.Count + 1000, 1).Address, "$", "")), "g"
'''                                KillColImPaire = False
'''                            Else
'''                                KillColImPaire = True
'''                            End If
'''                         End If
'''                     Next
''               End If
'    If CrerSoutotal(SheetFiltreCible, Id_Onglet, Id_IndiceProjet, Fils, NbDoc) = False Then
'        DeletSheet SheetFiltreCible
'    End If
'
''FinRd
'   RsFiltre.MoveNext
'   Wend
'End Sub
