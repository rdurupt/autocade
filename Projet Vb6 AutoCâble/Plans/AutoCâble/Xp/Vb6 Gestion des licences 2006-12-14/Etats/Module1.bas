Attribute VB_Name = "Module1"
Option Explicit
Global MyExcel As EXCEL.Application
Global CollecFiltreName As Collection

Type T_FiltreColonnes
    Cellule As New Collection
End Type
Type T_filtreLigne
    CollecLigne As New Collection
    Col As T_FiltreColonnes
End Type
Type T_Filtre
    Filtre As T_filtreLigne
End Type
Global MyFitre() As T_Filtre



'Function MyReplace(strVal As String) As String
'MyReplace = strVal
'MyReplace = Replace(MyReplace, "'", "''")
'MyReplace = Replace(MyReplace, Chr(34), Chr(34) & Chr(34))
'End Function
Function GenerEtatDapresDoc(Id_Macro As Long, Id_IndiceProjet As Long, Fils As Long, _
                            FileExcel As String, PDF As Boolean, Mydoc As GenerateurDoc, ColecDoc As Collection, NumDoc As Long) As Boolean
Dim Trie As String
Dim I As Long
Dim sql As String
Dim RsMacro As Recordset
Dim RsSelectChamp As Recordset
Dim RsExcel As Recordset
Dim Table As String
Dim MyWorkbook As EXCEL.Workbook
Dim MyWorkbookSource As EXCEL.Workbook
Dim MySheet As Worksheet
Dim MyRange As Range
Set MyExcel = New EXCEL.Application
Set MyWorkbook = MyExcel.Workbooks.Add
MyExcel.DisplayAlerts = False

MyExcel.Visible = True
For I = 1 To MyWorkbook.Worksheets.Count
MyWorkbook.Worksheets(I).Name = Replace(MyWorkbook.Worksheets(I).Name, "Feuil", "§Feuil§", 1)
Next
sql = "SELECT DISTINCT  T_ETATS.EtatName,T_ETATS.ID, T_Etats_Onglet.Id as Id_Select, "
sql = sql & "T_Etats_Onglet.Onglet, T_Etats_Onglet.Document ,T_Etats_Onglet.SaveOnglet, T_Menu_Etat_Onglet.TableName, T_Etats_Onglet.VueEpissur "
sql = sql & "FROM (T_ETATS INNER JOIN T_Etats_Onglet ON T_ETATS.ID = T_Etats_Onglet.Id_Etat) INNER JOIN T_Menu_Etat_Onglet ON T_Etats_Onglet.Onglet = T_Menu_Etat_Onglet.Menu "
sql = sql & "Where T_ETATS.Id = " & Id_Macro & " "
sql = sql & "ORDER BY T_Etats_Onglet.Id;"
Set RsMacro = Con.OpenRecordSet(sql)
BoolGenEtatEpisure = False

While RsMacro.EOF = False
If RsMacro!VueEpissur = True Then BoolGenEtatEpisure = True


If Trim("" & RsMacro!Document) = "" Then
    sql = "SELECT T_Etats_Select_Champs.ChamsName, T_Etats_Select_Champs.ChampAs, T_Etats_Select_Champs.Trie,T_Etats_Select_Champs.Visible "
    sql = sql & "From T_Etats_Select_Champs "
    sql = sql & "Where T_Etats_Select_Champs.Id_Onglet =" & RsMacro!Id_Select & " "
    sql = sql & "ORDER BY T_Etats_Select_Champs.Id;"

    
    Set RsSelectChamp = Con.OpenRecordSet(sql)
    If RsSelectChamp.EOF = False Then
            Table = "" & RsMacro!TableName
            sql = "Select "
            Trie = ""
        While RsSelectChamp.EOF = False
        If RsSelectChamp!Trie <> 0 Then
            Trie = Trie & "§" & RsSelectChamp!ChamsName & "=" & RsSelectChamp!Trie
        End If
            sql = sql & "[" & Table & "].[" & RsSelectChamp!ChamsName & "]"
            If UCase(RsSelectChamp!ChamsName) <> UCase(RsSelectChamp!ChampAs) Then
            sql = sql & " AS [" & RsSelectChamp!ChampAs & "]"
            End If
            sql = sql & ","
            RsSelectChamp.MoveNext
        Wend
        sql = Left(sql, Len(sql) - 1)
        sql = sql & " FROM " & Table & " Where " & Table & ".Id_IndiceProjet=" & Id_IndiceProjet & "; "
        
        Set RsExcel = Con.OpenRecordSet(sql)
        CreatEtaExels MyWorkbook, RsExcel, "" & RsMacro!SaveOnglet, Trie, RsMacro!Id_Select, Id_IndiceProjet, Fils, Id_Macro, NumDoc
        
        Set RsExcel = Con.CloseRecordSet(RsExcel)
     End If
     
    Set RsSelectChamp = Con.CloseRecordSet(RsSelectChamp)
   

Else
Set MyWorkbookSource = MyExcel.Workbooks.Open(ColecDoc(LstColecDoc(Replace("" & RsMacro!Document, " ", "_"))).SaveAs & ".XLS")
'Set MyRange = MyWorkBookSource.Worksheets(2).Range("A1").CurrentRegion
    sql = "SELECT T_Etats_Select_Champs.ChamsName, T_Etats_Select_Champs.ChampAs, T_Etats_Select_Champs.Trie,T_Etats_Select_Champs.Visible "
sql = sql & "From T_Etats_Select_Champs "
sql = sql & "Where T_Etats_Select_Champs.Id_Onglet =" & RsMacro!Id_Select & " "
sql = sql & "ORDER BY T_Etats_Select_Champs.Id;"

    
    Set RsSelectChamp = Con.OpenRecordSet(sql)
    If RsSelectChamp.EOF = False Then
'        Select Case RsMacro!Onglet
'                Case "Critères"
'                    Table = "T_Critères"
'                Case "Connecteurs"
'                    Table = "Connecteurs"
'                Case "Tableau de fils"
'                    Table = "Ligne_Tableau_fils"
'                Case "Composants"
'                    Table = "Composants"
'                Case "Notas"
'                    Table = "Nota"
'                Case "Noeuds"
'                    Table = "T_Noeuds"
'        End Select
Table = "" & RsMacro!TableName
            sql = "Select "
            Trie = ""
        While RsSelectChamp.EOF = False
        If RsSelectChamp!Trie <> 0 Then
            Trie = Trie & "§" & RsSelectChamp!ChamsName & "=" & RsSelectChamp!Trie
        End If
        
'           For i = 1 To MyRange.Columns.Count
'                If UCase(MyRange(1, i)) = UCase(RsSelectChamp!ChamsName) Then
'                    MyRange(1, i) = RsSelectChamp!ChampAs
'                    MyRange(1, i).Interior.ColorIndex = 2
'                    Exit For
'                End If
'           Next
            RsSelectChamp.MoveNext
        Wend
'       For i = 1 To MyRange.Columns.Count
'        If MyRange(1, i).Interior.ColorIndex <> 2 Then
'        DelColonne MyWorkBookSource.Worksheets(2), i
'        End If
'       Next
       
       
'        Set RsExcel = Con.OpenRecordSet(Sql)
        CreatEtaExelsDapresDoc MyWorkbook, MyWorkbookSource, "" & RsMacro!SaveOnglet, Trie, RsMacro!Id_Select, Id_IndiceProjet, Fils, Id_Macro, FileExcel, RsSelectChamp, NumDoc
        MyWorkbookSource.Close False
'        Set RsExcel = Con.CloseRecordSet(RsExcel)
     End If
     
    Set RsSelectChamp = Con.CloseRecordSet(RsSelectChamp)
End If
    RsMacro.MoveNext
Wend
DeletSheetEtat MyWorkbook, ActionType, 0
Set MySheet = IsertSheet(MyWorkbook, "MyFilterElaborré")
MySheet.Application.DisplayAlerts = False
MySheet.Delete
If PDF = True Then
 PrintPdf MyWorkbook, FileExcel & ".PDF"
Else
    MyWorkbook.SaveAs FileExcel
End If
MyWorkbook.Close False
MyExcel.Quit
Set MyWorkbook = Nothing
Set MyExcel = Nothing
GenerEtatDapresDoc = True
End Function
Public Function GenerEtat(Id_Macro As Long, Id_IndiceProjet As Long, Fils As Long, FileExcel As String, PDF As Boolean, NbDoc As Long) As Boolean
Dim Trie As String
Dim I As Long
Dim sql As String
Dim RsMacro As Recordset
Dim RsSelectChamp As Recordset
Dim RsExcel As Recordset
Dim Table As String
Dim MyWorkbook As EXCEL.Workbook
Dim MySheet As Worksheet
Set MyExcel = New EXCEL.Application
Set MyWorkbook = MyExcel.Workbooks.Add
Dim SqlChamp As String

MyExcel.Visible = True
MyExcel.DisplayAlerts = False
For I = 1 To MyWorkbook.Worksheets.Count
MyWorkbook.Worksheets(I).Name = Replace(MyWorkbook.Worksheets(I).Name, "Feuil", "§Feuil§", 1)
Next
sql = "SELECT DISTINCT  T_ETATS.EtatName,T_ETATS.ID, T_Etats_Onglet.Id as Id_Select, "
sql = sql & "T_Etats_Onglet.Onglet, T_Etats_Onglet.Document ,T_Etats_Onglet.SaveOnglet, T_Menu_Etat_Onglet.TableName,T_Etats_Onglet.VueEpissur "
sql = sql & "FROM (T_ETATS INNER JOIN T_Etats_Onglet ON T_ETATS.ID = T_Etats_Onglet.Id_Etat) INNER JOIN T_Menu_Etat_Onglet ON T_Etats_Onglet.Onglet = T_Menu_Etat_Onglet.Menu "
sql = sql & "Where T_ETATS.Id = " & Id_Macro & " "
sql = sql & "ORDER BY T_Etats_Onglet.Id;"

BoolGenEtatEpisure = False

    
Set RsMacro = Con.OpenRecordSet(sql)
While RsMacro.EOF = False
If RsMacro!VueEpissur = True Then BoolGenEtatEpisure = True
    sql = "SELECT T_Etats_Select_Champs.ChamsName, T_Etats_Select_Champs.ChampAs, T_Etats_Select_Champs.Trie,T_Etats_Select_Champs.Visible "
sql = sql & "From T_Etats_Select_Champs "
sql = sql & "Where T_Etats_Select_Champs.Id_Onglet =" & RsMacro!Id_Select & " "
sql = sql & "ORDER BY T_Etats_Select_Champs.Id;"

    
    Set RsSelectChamp = Con.OpenRecordSet(sql)
    If RsSelectChamp.EOF = False Then
'        Select Case RsMacro!Onglet
'                Case "Critères"
'                    Table = "T_Critères"
'                Case "Connecteurs"
'                    Table = "Connecteurs"
'                Case "Tableau de fils"
'                    Table = "Ligne_Tableau_fils"
'                Case "Composants"
'                    Table = "Composants"
'                Case "Notas"
'                    Table = "Nota"
'                Case "Noeuds"
'                    Table = "T_Noeuds"
'        End Select
        Table = "" & RsMacro!TableName
            sql = "Select "
            Trie = ""
        While RsSelectChamp.EOF = False
        If RsSelectChamp!Trie <> 0 Then
            Trie = Trie & "§" & RsSelectChamp!ChamsName & "=" & RsSelectChamp!Trie
        End If
            sql = sql & "[" & Table & "].[" & RsSelectChamp!ChamsName & "]"
            If UCase(RsSelectChamp!ChamsName) <> UCase(RsSelectChamp!ChampAs) Then
            sql = sql & " AS [" & RsSelectChamp!ChampAs & "]"
            End If
            sql = sql & ","
            RsSelectChamp.MoveNext
        Wend
        sql = Left(sql, Len(sql) - 1)
        sql = sql & " FROM " & Table & " Where " & Table & ".Id_IndiceProjet=" & Id_IndiceProjet & "; "
        
        SqlChamp = sql
        sql = "select [" & Table & "].*  FROM " & Table & " Where " & Table & ".Id_IndiceProjet=" & Id_IndiceProjet & "; "
        Set RsExcel = Con.OpenRecordSet(sql)
        
        
        If NbDoc = 0 Then
            CreatEtaExels MyWorkbook, RsExcel, "" & RsMacro!SaveOnglet, Trie, RsMacro!Id_Select, Id_IndiceProjet, Fils, Id_Macro, NbDoc
        Else
            CreatEtaExels MyWorkbook, RsExcel, "" & RsMacro!SaveOnglet, Trie, RsMacro!Id_Select, Id_IndiceProjet, Fils, Id_Macro, NbDoc, True
        End If
        Set RsExcel = Con.CloseRecordSet(RsExcel)
     End If
     
    Set RsSelectChamp = Con.CloseRecordSet(RsSelectChamp)
    RsMacro.MoveNext
Wend
DeletSheetEtat MyWorkbook, ActionType, 0
Set MySheet = IsertSheet(MyWorkbook, "MyFilterElaborré")
MySheet.Application.DisplayAlerts = False
MySheet.Delete
If PDF = True Then
 PrintPdf MyWorkbook, FileExcel & ".PDF"
Else
    MyWorkbook.SaveAs FileExcel
End If
MyWorkbook.Close False
MyExcel.Quit
Set MyWorkbook = Nothing
Set MyExcel = Nothing
GenerEtat = True
End Function
Function CrerSoutotal(MySheet As Worksheet, Id_Onglet As Long, Id_IndiceProjet As Long, Fils As Long, NbDoc As Long, Optional BoolFill As Boolean = True, Optional Executer As Boolean = True, Optional OfseteMiseEnPage As Long, Optional EnteteOnglet As String) As Boolean
Dim MyRange As Range
Dim txtOnglet As String
Dim Rs As Recordset
Dim Index As Long
Dim C As Long
Dim L As Long
Dim R1 As Long
Dim R2 As Long
Dim RsOnlget As Recordset
Dim RsEntetePage As Recordset
Dim IdEntete As Long
Dim sql As String

If Fils <> 0 Then
IdEntete = Fils
Else
    IdEntete = Id_IndiceProjet
End If
   
            Set MyRange = MySheet.Range("A1").CurrentRegion
            Debug.Print MySheet.Cells(1, 1).Address
            
            Set MyRange = MySheet.Range("A1:" & Replace(MySheet.Cells(MyRange.Rows.Count + OfseteMiseEnPage, MyRange.Columns.Count).Address, "$", ""))
            
             sql = "SELECT  T_indiceProjet.Ensemble, T_indiceProjet.Equipement,[Li] & '_' & [LI_Indice] AS Liste, "
             sql = sql & "[PI] & '_' & [PI_Indice] AS Piece, T_indiceProjet.Client, T_indiceProjet.CleAc "
             sql = sql & "FROM T_indiceProjet "
             sql = sql & "WHERE T_indiceProjet.Id=" & IdEntete & " ;"
              Set RsEntetePage = Con.OpenRecordSet(sql)
             txtOnglet = ConverOngletGneEtat(MySheet.Name)
             
             If Trim("" & EnteteOnglet) <> "" Then txtOnglet = EnteteOnglet
             If PerfEntete = False Then
                sql = "SELECT  T_Etats_Onglet.Macro,T_Etats_Onglet.Onglet FROM T_Etats_Onglet GROUP BY T_Etats_Onglet.Onglet, T_Etats_Onglet.Macro;"
                Set RsOnlget = Con.OpenRecordSet(sql)
                While RsOnlget.EOF = False
                    txtOnglet = Replace(UCase(txtOnglet), UCase(Trim("" & RsOnlget!Macro)) & "_", "")
                    RsOnlget.MoveNext
                Wend
             End If
             txtOnglet = "&14&""Arial,Gras""" & txtOnglet & "&10&""Arial,Normal"""
            If PortraitPaysage = 0 Then PortraitPaysage = 2
             MiseEnPage MySheet, MyRange, "Affaire : " & RsEntetePage!CleAc & Chr(10) & _
             "Pièce : " & RsEntetePage!Piece & Chr(10) & "Liste : " & RsEntetePage!Liste, _
                 txtOnglet & Chr(10) & "&10Câblage : " & Replace("" & RsEntetePage!Ensemble, vbCrLf, " ") & Chr(10) & "Equipement : " & Replace("" & RsEntetePage!Equipement, vbCrLf, " ") & Chr(10) & _
             "" _
             , "ENCELADE" & Chr(10) & "Client : " & RsEntetePage!Client & Chr(10) & Format(Date, "dd-mmm-yyyy"), "", "&P/&N", "", 80, "C2", True, PortraitPaysage, True

            MaJEncadreXls MySheet.Range("A1").CurrentRegion, xlThin, xlThin, xlHairline, xlHairline

             Set RsEntetePage = Con.CloseRecordSet(RsEntetePage)
If MySheet.Range("a1").CurrentRegion.Rows.Count < 2 Then Exit Function
             
If Executer = False Then GoTo Fin
If NbDoc <> 0 Then
If BoolFill = False Then CrerSoutotal = True
Exit Function
End If
sql = "SELECT T_Etats_Select_Champs.ChampAs FROM T_Etats_Select_Champs "
sql = sql & "WHERE T_Etats_Select_Champs.SousTautal=True AND T_Etats_Select_Champs.Id_Onglet=" & Id_Onglet & ";"
Set Rs = Con.OpenRecordSet(sql)
If Rs.EOF = False Then
     For Index = 1 To 4
        InsertLIgneExcel MySheet, 1
     Next
     Set MyRange = MySheet.Cells(5, 1).CurrentRegion
     While Rs.EOF = False
       
            For C = 1 To MyRange.Columns.Count
                If UCase(MyRange(1, C)) = UCase(Rs(0)) Then
                R1 = MySheet.Range(MyRange(2, MyRange.Columns.Count).Address).Row
                R2 = MySheet.Range(MyRange(MyRange.Rows.Count, MyRange.Columns.Count).Address).Row
                MySheet.Cells(2, C) = "Sous Total"
                MySheet.Cells(3, C).FormulaR1C1 = "=SUBTOTAL(9,R[" & R1 - 3 & "]C:R[" & R2 - 3 & "]C)"
                
                FormatExcelPlage MySheet.Cells(2, C), 15, False, True, xlCenter, xlCenter
                FormatExcelPlage MySheet.Cells(3, C), 2, False, True, xlCenter, xlCenter
                Exit For
                End If
            Next
        
        Rs.MoveNext
     Wend
End If
Fin:
CrerSoutotal = True
End Function




Function CreatEtaExelsDapresDoc(MyExel As EXCEL.Workbook, MyWorkbookSource As Workbook, Onglet As String, Trie As String, Id_Onglet As Long, Id_IndiceProjet As Long, Fils As Long, Id_Macro As Long, Document As String, RsSelectChamp As Recordset, NbDoc As Long) As Boolean
Dim DapresDoc As Boolean
DapresDoc = False
If NbDoc <> 0 Then DapresDoc = True
Dim MySeet As EXCEL.Worksheet
Dim MyRange As EXCEL.Range
Dim SheetFiltre As Worksheet
Dim SheetFiltreCible As Worksheet
Dim ColOption As Long
Dim SaveNbL As Long
'Set MySeet = IsertSheet(MyExel, Onglet)
'Set MyRange = MySeet.Range("A1").CurrentRegion
Dim RsCriteres As Recordset
Dim RsOptionOnglet As Recordset
Dim IndexOnglet As Long
Dim MyTrie
Dim MyTrieS
Dim sql As String
Dim RsFiltre As Recordset
Dim RsFiltreELab As Recordset
Dim I As Long
Dim DebutOnglet As Long
Dim FinOnget As Long
Dim TrieExcel(3, 1) As String
Dim Col As Long
Dim NbTrie As Integer
Dim IsFilterElab As Boolean
Dim MyOption
Dim NbMyOption As Long
Dim RangeCopy As Range
Dim ClsDog As GenerateurDoc
Dim RsChampMasquer As Recordset
Dim PoseFiterMyOption As Long
Dim Debut As Long
Dim I2, I3, I4, I5, I6, I7, I8 As Long
Dim L, C As Long
Dim SplitLstOnglet
Dim SplitEpisure
Dim OpTionValue
Dim PrefixName As String
Dim KillColImPaire As Boolean
 sql = "SELECT DISTINCT T_Etats_Onglet.Id AS Id_Select, T_Etats_Select_Champs.ChamsName, T_Etats_Select_Champs.Visible "
    sql = sql & "FROM T_Etats_Onglet INNER JOIN T_Etats_Select_Champs ON T_Etats_Onglet.Id = T_Etats_Select_Champs.Id_Onglet "
    sql = sql & "Where T_Etats_Onglet.Id = " & Id_Onglet & " "
    sql = sql & "And T_Etats_Select_Champs.Visible = false; "
    Set RsChampMasquer = Con.OpenRecordSet(sql)
'ExcelCreatTitre MyRange, rs
'While rs.EOF = False
'    Set MyRange = MySeet.Range("A1").CurrentRegion
'    Set MyRange = MySeet.Range(MyRange(MyRange.Rows.Count + 1, 1).Address & ":" & MyRange(MyRange.Rows.Count + 1, MyRange.Columns.Count).Address)
'    ExcelCreatTitre MyRange, rs, True
'    rs.MoveNext
' Wend
 MyTrie = Split(Trie & "§", "§")
 If UBound(MyTrie) > 1 Then
    For I = 1 To UBound(MyTrie)
        MyTrieS = Split(MyTrie(I) & "=", "=")
    If Val(MyTrieS(1)) <> 0 Then
    TrieExcel(I, 0) = MyTrieS(0): TrieExcel(I, 1) = MyTrieS(1)
    If I = 3 Then Exit For
    End If
   
 Next
 
' Set MyRange = MySeet.Range("A1").CurrentRegion
'NbTrie = 0
' For i = 1 To 3
'    For Col = 1 To MyRange.Columns.Count
'       If MyRange(1, Col) = TrieExcel(i, 0) Then
'       NbTrie = NbTrie + 1
'           TrieExcel(i, 0) = MyRange(2, Col).Address
'            TrieExcel(i, 0) = Replace(TrieExcel(i, 0), "$", "")
'           Exit For
'       End If
'    Next
'  Next
 End If
 IndexOnglet = 0
sql = "SELECT T_Etats_Onglet.* From T_Etats_Onglet WHERE T_Etats_Onglet.Id=" & Id_Onglet & ";"
Set RsOptionOnglet = Con.OpenRecordSet(sql)
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
            If DebutOnglet = 0 Then Exit Function
        End If
    End If
     DebutOnglet = DebutOnglet + Val("" & RsOptionOnglet!DecaleAppres)
     If DebutOnglet < 1 And DebutOnglet > MyWorkbookSource.Sheets.Count Then Exit Function
     
     
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
            If FinOnget = 0 Then Exit Function
        End If
    End If
     FinOnget = FinOnget - Val("" & RsOptionOnglet!DecaleAvant)
     If FinOnget < 1 And FinOnget > MyWorkbookSource.Sheets.Count Then Exit Function
    If DebutOnglet > FinOnget Then Exit Function
    For I = FinOnget To DebutOnglet Step -1
'        If Left(MyWorkBookSource.Sheets(I).Name, Len("" & RsOptionOnglet!OngletStrat)) = "" & RsOptionOnglet!OngletStrat Then
            
            Set MySeet = IsertSheet(MyExel, Onglet & "_" & MyWorkbookSource.Sheets(I).Name)
'            MyWorkBookSource.Activate
'            MyWorkBookSource.Sheets(I).Select
            MyWorkbookSource.Sheets(I).Range("A1").CurrentRegion.Copy
            MyExel.Activate
            MySeet.Select
            MySeet.Range("A1").Select
            MySeet.Paste
             Set ClsDog = New GenerateurDoc
            ClsDog.RenseigneColonn MySeet, Id_Onglet, 1
            RsSelectChamp.Requery
            Set MyRange = MySeet.Range("A1").CurrentRegion
            If Trim(ClsDog.LstOngletName) = "" Then
                         For I2 = 1 To MyRange.Columns.Count
            If MyRange(1, I2) = "OPTION" Then Exit For
         Next
          For I3 = 2 To MyRange.Rows.Count
            If Trim("" & MyRange(I3, I2)) = "" Then MyRange(I3, I2) = "§Null§"
        Next
         Set MyRange = MySeet.Range("A1").CurrentRegion
          Trier MySeet, NbTrie, Replace("A2:" & MyRange(MyRange.Rows.Count, MyRange.Columns.Count).Address, "$", ""), TrieExcel(1, 0), Val(TrieExcel(1, 1)), TrieExcel(2, 0), Val(TrieExcel(2, 1)), TrieExcel(3, 0), Val(TrieExcel(3, 1))

            Set CollecFiltreName = Nothing
            Set CollecFiltreName = New Collection
            sql = "SELECT T_Etats_Select_Filtre.FiltreName "
            sql = sql & "From T_Etats_Select_Filtre "
            sql = sql & "Where T_Etats_Select_Filtre.Id_Onglet =  " & Id_Onglet & " "
            sql = sql & "GROUP BY T_Etats_Select_Filtre.FiltreName "
            sql = sql & "ORDER BY T_Etats_Select_Filtre.FiltreName; "
            
            
            Set RsFiltre = Con.OpenRecordSet(sql)
            I3 = 0
            If RsFiltre.EOF = True Then
            
            sql = "SELECT T_Etats_Select_Champs.ChampAs, T_indiceProjet.Equipement, T_Etats_Select_Champs.Id_Onglet "
            sql = sql & "FROM T_indiceProjet, T_Etats_Onglet INNER JOIN T_Etats_Select_Champs ON T_Etats_Onglet.Id = T_Etats_Select_Champs.Id_Onglet "
            sql = sql & "WHERE (T_Etats_Select_Champs.ChamsName='OPTION' OR T_Etats_Select_Champs.ChamsName='CRITERES') "
            If Fils = 0 Then
                sql = sql & "AND T_indiceProjet.Id=" & Id_IndiceProjet & " "
            Else
                sql = sql & "AND T_indiceProjet.Id=" & Fils & " "
            End If
            sql = sql & "AND T_Etats_Select_Champs.Id_Onglet=" & Id_Onglet & " AND T_Etats_Onglet.FiltreEquipement=True "
            sql = sql & "GROUP BY T_Etats_Select_Champs.ChamsName, T_Etats_Select_Champs.ChampAs, T_indiceProjet.Equipement, T_indiceProjet.Id, "
            sql = sql & "T_Etats_Select_Champs.Id_Onglet;"
            
            
            Set RsFiltreELab = Con.OpenRecordSet(sql)
            If RsFiltreELab.EOF = False Then
            IsFilterElab = True
            MyOption = FunCodeCritaire(Id_IndiceProjet, "" & RsFiltreELab!Equipement)
                NbMyOption = 0
            '    MyOption = Split("" & RsFiltreELab!Equipement & ";", ";")
            
                Set RangeCopy = SheetFiltre.Range("a1").CurrentRegion
                For I3 = 0 To UBound(MyOption)
                    PoseFiterMyOption = SheetFiltre.Range("a1").CurrentRegion.Rows.Count
                    If Trim("" & MyOption(I3)) <> "" Then
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
                        For I4 = Debut To SheetFiltre.Range("a1").CurrentRegion.Rows.Count
                            For I5 = 1 To SheetFiltre.Range("a1").CurrentRegion.Columns.Count
                                If SheetFiltre.Cells(1, I5) = RsFiltreELab!ChampAs Then
                                    SheetFiltre.Cells(I4, I5) = "*;" & Replace("" & MyOption(I3), "§Null§", "") & ";*"
'                                    SheetFiltre.Cells(I4, I5) = MyOption(I3)
                                    Exit For
                                End If
                            Next
                        Next
                        NbMyOption = NbMyOption + 1
                End If
            Next
            For C = 1 To MyRange.Columns.Count
                If UCase(MyRange(1, C)) = "OPTION" Or UCase(MyRange(1, C)) = "CRITERES" Then
                    For L = 2 To MyRange.Rows.Count
                        MyRange(L, C) = ";" & MyRange(L, C) & ";"
                    Next
                    Exit For
                End If
             Next
            Set SheetFiltreCible = IsertSheet(MyExel, MySeet.Name & "_FLT_Equipement")
            FiltreActif MySeet.Range("a1").CurrentRegion, SheetFiltre.Range("A1").CurrentRegion, SheetFiltreCible.Range("a1").CurrentRegion
            
             Set MyRange = SheetFiltreCible.Range("a1").CurrentRegion
            For C = 1 To MyRange.Columns.Count
                If UCase(MyRange(1, C)) = "OPTION" Or UCase(MyRange(1, C)) = "CRITERES" Then
                    For L = 2 To MyRange.Rows.Count
                        MyRange(L, C) = Left(MyRange(L, C), Len(MyRange(L, C)) - 1)
                        MyRange(L, C) = Right(MyRange(L, C), Len(MyRange(L, C)) - 1)
                    Next
                    Exit For
                End If
            Next
            If BoolGenEtatEpisure = True Then
               
                VueArriere MySeet
               End If
              CrerSoutotal SheetFiltreCible, Id_Onglet, Id_IndiceProjet, Fils, NbDoc, False
              Else
            If BoolGenEtatEpisure = True Then
                SplitEpisure = Split(MySeet.Name & "E", "E")
                VueArriere MySeet
               End If
              CrerSoutotal MySeet, Id_Onglet, Id_IndiceProjet, Fils, NbDoc, False
            End If
            End If
            
            
            While RsFiltre.EOF = False
            IsFilterElab = True
            Set SheetFiltre = IsertSheet(MyExel, "MyFilterElaborré")
            DeleteRow SheetFiltre, True
            Set MyRange = MySeet.Range("A1").CurrentRegion
            For I3 = 1 To MyRange.Columns.Count
            SheetFiltre.Cells(1, I3) = MyRange(1, I3)
            sql = "SELECT T_Etats_Select_Filtre.FiltreName, T_Etats_Select_Champs.ChampAs, T_Etats_Select_Filtre.Valeur, T_Etats_Select_Filtre.Ligne "
            sql = sql & "FROM (T_Etats_Onglet INNER JOIN T_Etats_Select_Filtre ON T_Etats_Onglet.Id = T_Etats_Select_Filtre.Id_Onglet)  "
            sql = sql & "INNER JOIN T_Etats_Select_Champs ON (T_Etats_Select_Filtre.Colonne = T_Etats_Select_Champs.ChamsName)  "
            sql = sql & "AND (T_Etats_Select_Filtre.Id_Onglet = T_Etats_Select_Champs.Id_Onglet) "
            sql = sql & "WHERE T_Etats_Select_Champs.ChampAs='" & MyReplace(MyRange(1, I3)) & "' "
            sql = sql & "AND T_Etats_Select_Filtre.Id_Onglet=" & Id_Onglet & " AND T_Etats_Select_Filtre.FiltreName='" & MyReplace(RsFiltre!FiltreName) & "' "
            sql = sql & "ORDER BY T_Etats_Select_Filtre.Id;"
            
            Set RsFiltreELab = Con.OpenRecordSet(sql)
            While RsFiltreELab.EOF = False
            
            
            
            SheetFiltre.Cells(RsFiltreELab!Ligne, I3) = "" & RsFiltreELab!Valeur
            SaveNbL = RsFiltreELab!Ligne
            
            
            RsFiltreELab.MoveNext
            Wend
            Next
            
            Set RsFiltreELab = Con.CloseRecordSet(RsFiltreELab)
            For I3 = 2 To SaveNbL
            If SheetFiltre.Range("a1").CurrentRegion.Rows.Count = 1 Then
            SheetFiltre.Rows(CStr(2) & ":" & CStr(2)).Delete Shift:=xlUp
            
            End If
            Next
            sql = "SELECT T_Etats_Select_Champs.ChampAs, T_indiceProjet.Equipement, T_Etats_Select_Champs.Id_Onglet "
            sql = sql & "FROM T_indiceProjet, T_Etats_Onglet INNER JOIN T_Etats_Select_Champs ON T_Etats_Onglet.Id = T_Etats_Select_Champs.Id_Onglet "
            sql = sql & "WHERE (T_Etats_Select_Champs.ChamsName='OPTION' OR T_Etats_Select_Champs.ChamsName='CRITERES') "
            If Fils = 0 Then
                sql = sql & "AND T_indiceProjet.Id=" & Id_IndiceProjet & " "
            Else
                sql = sql & "AND T_indiceProjet.Id=" & Fils & " "
            End If
            sql = sql & "AND T_Etats_Select_Champs.Id_Onglet=" & Id_Onglet & " AND T_Etats_Onglet.FiltreEquipement=True "
            sql = sql & "GROUP BY T_Etats_Select_Champs.ChamsName, T_Etats_Select_Champs.ChampAs, T_indiceProjet.Equipement, T_indiceProjet.Id, "
            sql = sql & "T_Etats_Select_Champs.Id_Onglet;"
            
            
            Set RsFiltreELab = Con.OpenRecordSet(sql)
            If RsFiltreELab.EOF = False Then
                IsFilterElab = True
            MyOption = FunCodeCritaire(Id_IndiceProjet, "" & RsFiltreELab!Equipement)
                NbMyOption = 0
            '    MyOption = Split("" & RsFiltreELab!Equipement & ";", ";")
            
                Set RangeCopy = SheetFiltre.Range("a1").CurrentRegion
                For I3 = 0 To UBound(MyOption)
                    PoseFiterMyOption = SheetFiltre.Range("a1").CurrentRegion.Rows.Count
                    If Trim("" & MyOption(I3)) <> "" Then
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
                        For I4 = Debut To SheetFiltre.Range("a1").CurrentRegion.Rows.Count
                            For I5 = 1 To SheetFiltre.Range("a1").CurrentRegion.Columns.Count
                                If SheetFiltre.Cells(1, I5) = RsFiltreELab!ChampAs Then
                                    SheetFiltre.Cells(I4, I5) = "*;" & Replace("" & MyOption(I3), "§Null§", "") & ";*"
'                                    SheetFiltre.Cells(I4, I5) = MyOption(I3)
                                    Exit For
                                End If
                            Next
                        Next
                        NbMyOption = NbMyOption + 1
                End If
            Next
            End If
           
            Set MyRange = MySeet.Range("a1").CurrentRegion
            For C = 1 To MyRange.Columns.Count
                If UCase(MyRange(1, C)) = "OPTION" Or UCase(MyRange(1, C)) = "CRITERES" Then
                    For L = 2 To MyRange.Rows.Count
                        MyRange(L, C) = ";" & MyRange(L, C) & ";"
                    Next
                    Exit For
                End If
             Next
            Set SheetFiltreCible = IsertSheet(MyExel, MySeet.Name & "_FLT_Equipement")
            FiltreActif MySeet.Range("a1").CurrentRegion, SheetFiltre.Range("A1").CurrentRegion, SheetFiltreCible.Range("a1").CurrentRegion
            
            Set MyRange = MySeet.Range("A1").CurrentRegion
           
             RsSelectChamp.Requery
            RsSelectChamp.Filter = "Visible=false"
            RsChampMasquer.Requery
            While RsChampMasquer.EOF = False
            For I2 = 1 To MyRange.Columns.Count
                If UCase(MyRange(1, I2)) = UCase(RsSelectChamp!ChamsName) Then
                    MyRange(1, I2) = RsSelectChamp!ChampAs
                    MyRange(1, I2).Interior.ColorIndex = 6
                    Exit For
                End If
           Next
            RsChampMasquer.MoveNext
        Wend
        
         If BoolGenEtatEpisure = True Then
               
                VueArriere MySeet
               End If
          Set MyRange = MySeet.Range("a1").CurrentRegion
            For C = 1 To MyRange.Columns.Count
                If UCase(MyRange(1, C)) = "OPTION" Or UCase(MyRange(1, C)) = "CRITERES" Then
                    For L = 2 To MyRange.Rows.Count
                        MyRange(L, C) = Left(MyRange(L, C), Len(MyRange(L, C)) - 1)
                        MyRange(L, C) = Right(MyRange(L, C), Len(MyRange(L, C)) - 1)
                    Next
                    Exit For
                End If
            Next
            Set MyRange = MySeet.Range("a1").CurrentRegion
            
            If DapresDoc = False Then
                    For I5 = MySeet.Range("A1").CurrentRegion.Columns.Count To 1 Step -1
                        If MySeet.Cells(1, I5).Interior.ColorIndex <> 6 Then
                            SuprmerCells MySeet.Range(Replace(MySeet.Cells(1, I5).Address & ":" & MySeet.Cells(MySeet.Range("A1").CurrentRegion.Rows.Count, I5).Address, "$", "")), "g"
                            If KillColImPaire = True Then
                                SuprmerCells MySeet.Range(Replace(MySeet.Cells(MySeet.Range("A1").CurrentRegion.Rows.Count + 1, 1).Address & ":" & MySeet.Cells(MySeet.Range("A1").CurrentRegion.Rows.Count + 1000, 1).Address, "$", "")), "g"
                                KillColImPaire = False
                            Else
                                KillColImPaire = True
                            End If
                         End If
                     Next
                End If
'            For I2 = Myrange.Columns.Count To 1 Step -1
'            If Myrange(1, I2).Interior.ColorIndex <> 6 Then
'                DelColonne SheetFiltreCible, Val(I2)
'            End If
       
            CrerSoutotal MySeet, Id_Onglet, Id_IndiceProjet, Fils, NbDoc
            
            
            
            
            
            RsFiltre.MoveNext
            
            Wend
'            MySeet.Application.DisplayAlerts = False
'             If BoolGenEtatEpisure = True Then
'
'                VueArriere MySeet, "" & SplitEpisure(0)
'               End If
'              CrerSoutotal MySeet, Id_Onglet, Id_IndiceProjet, Fils, NbDoc, False
'            If IsFilterElab = True Then MySeet.Delete
'            Else
'                SplitLstOnglet = Split(ClsDog.LstOnglet & ";", ";")
'                For I7 = 2 To MySeet.Range("a1").CurrentRegion.Rows.Count
'                    MySeet.Cells(I7, ClsDog.MyClonne(ClsDog.LstOngletName)) = ";" & MySeet.Cells(I7, ClsDog.MyClonne(ClsDog.LstOngletName)) & ";"
'                Next
'            For I4 = 0 To UBound(SplitLstOnglet)
'            If Trim(SplitLstOnglet(I4)) <> "" Then
'            Set SheetFiltre = IsertSheet(MyExel, "MyFilterElaborré")
'            If SheetFiltre.Range("A1").CurrentRegion.Rows.Count = 1 And SheetFiltre.Range("A1").CurrentRegion.Columns.Count = 1 Then
'                If Trim("" & SheetFiltre.Range("A1")) = "" Then
'                    PrefixName = "_Ongt_"
'                    For C = 1 To MySeet.Range("a1").CurrentRegion.Columns.Count
'                        SheetFiltre.Cells(1, C) = MySeet.Cells(1, C)
'                        SheetFiltre.Cells(2, ClsDog.MyClonne(ClsDog.LstOngletName)) = ";"
'                    Next
'                End If
'             End If
'                For I7 = 2 To SheetFiltre.Range("a1").CurrentRegion.Rows.Count
'                    SheetFiltre.Cells(I7, ClsDog.MyClonne(ClsDog.LstOngletName)) = ";" & SplitLstOnglet(I4) & ";"
'                Next
'                 Set SheetFiltreCible = IsertSheet(MyExel, MySeet.Name & "_FLT_S_Equ_" & SplitLstOnglet(I4), True)
'                FiltreActif MySeet.Range("a1").CurrentRegion, SheetFiltre.Range("A1").CurrentRegion, SheetFiltreCible.Range("a1").CurrentRegion
'
'                For I7 = 2 To SheetFiltreCible.Range("a1").CurrentRegion.Rows.Count
'                    SheetFiltreCible.Cells(I7, ClsDog.MyClonne(ClsDog.LstOngletName)) = Replace(SheetFiltreCible.Cells(I7, ClsDog.MyClonne(ClsDog.LstOngletName)), ";", "")
'                Next
'                If BoolGenEtatEpisure = True Then
'                    VueArriere SheetFiltreCible, MySeet.Name & "_FLT_S_Equ_"
'                End If
'                    RsChampMasquer.Requery
'                    KillColImPaire = False
'                While RsChampMasquer.EOF = False
''                SheetFiltreCible.Range("A1").Interior.ColorIndex
'                    SheetFiltreCible.Range(Replace(SheetFiltreCible.Cells(1, ClsDog.MyClonne("" & RsChampMasquer!ChamsName)).Address, "$", "")).Interior.ColorIndex = 6
'                     RsChampMasquer.MoveNext
'                Wend
'                If DapresDoc = False Then
'                    For I5 = SheetFiltreCible.Range("A1").CurrentRegion.Columns.Count To 1 Step -1
'                        If SheetFiltreCible.Cells(1, I5).Interior.ColorIndex <> 6 Then
'                            SuprmerCells SheetFiltreCible.Range(Replace(SheetFiltreCible.Cells(1, I5).Address & ":" & SheetFiltreCible.Cells(SheetFiltreCible.Range("A1").CurrentRegion.Rows.Count, I5).Address, "$", "")), "g"
'                            If KillColImPaire = True Then
'                                SuprmerCells SheetFiltreCible.Range(Replace(SheetFiltreCible.Cells(SheetFiltreCible.Range("A1").CurrentRegion.Rows.Count + 1, 1).Address & ":" & SheetFiltreCible.Cells(SheetFiltreCible.Range("A1").CurrentRegion.Rows.Count + 1000, 1).Address, "$", "")), "g"
'                                KillColImPaire = False
'                            Else
'                                KillColImPaire = True
'                            End If
'                         End If
'                     Next
'                End If
''    b
'            If BoolGenEtatEpisure = True Then
'                SplitEpisure = Split(MySeet.Name & "E", "E")
'                VueArriere MySeet, "" & SplitEpisure(0)
'               End If
'                If CrerSoutotal(MySeet, Id_Onglet, Id_IndiceProjet, Fils, NbDoc, False) = False Then
'                    DeletSheet MySeet
'                End If
'            End If
'            Next
'
'    End If
'         a

        End If
        
    Next
     
End If
If Trim("" & RsOptionOnglet!OngletStrat) <> "" And Trim("" & RsOptionOnglet!OngleEnd) = "" And RsOptionOnglet!FiltreSequentielle = True Then
    For I = 1 To MyWorkbookSource.Sheets.Count
        If Left(MyWorkbookSource.Sheets(I).Name, Len("" & RsOptionOnglet!OngletStrat)) = "" & RsOptionOnglet!OngletStrat Then
            IndexOnglet = IndexOnglet + 1
            Set MySeet = IsertSheet(MyExel, Onglet & "_" & RsOptionOnglet!OngletStrat & "_" & CStr(IndexOnglet))
            MyWorkbookSource.Sheets(I).Range("A1").CurrentRegion.Copy
            MyExel.Activate
            MySeet.Select
            MySeet.Range("A1").Select
            MySeet.Paste
            RsSelectChamp.Requery
            Set MyRange = MySeet.Range("A1").CurrentRegion
             RsSelectChamp.Requery
            RsSelectChamp.Filter = "Visible=false"
            While RsSelectChamp.EOF = False
            For I2 = 1 To MyRange.Columns.Count
                If UCase(MyRange(1, I2)) = UCase(RsSelectChamp!ChamsName) Then
                    MyRange(1, I2) = RsSelectChamp!ChampAs
                    MyRange(1, I2).Interior.ColorIndex = 6
                    Exit For
                End If
           Next
            RsSelectChamp.MoveNext
        Wend
        For I2 = MyRange.Columns.Count To 1 Step -1
            If MyRange(1, I2).Interior.ColorIndex <> 6 Then
                DelColonne MySeet, Val(I2)
            End If
         Next
         For I2 = 1 To MyRange.Columns.Count
            If MyRange(1, I2) = "OPTION" Then Exit For
         Next
          For I3 = 2 To MyRange.Rows.Count
            If Trim("" & MyRange(I3, I2)) = "" Then MyRange(I3, I2) = "§Null§"
        Next
         Set MyRange = MySeet.Range("A1").CurrentRegion
          Trier MySeet, NbTrie, Replace("A2:" & MyRange(MyRange.Rows.Count, MyRange.Columns.Count).Address, "$", ""), TrieExcel(1, 0), Val(TrieExcel(1, 1)), TrieExcel(2, 0), Val(TrieExcel(2, 1)), TrieExcel(3, 0), Val(TrieExcel(3, 1))

            Set CollecFiltreName = Nothing
            Set CollecFiltreName = New Collection
            sql = "SELECT T_Etats_Select_Filtre.FiltreName "
            sql = sql & "From T_Etats_Select_Filtre "
            sql = sql & "Where T_Etats_Select_Filtre.Id_Onglet =  " & Id_Onglet & " "
            sql = sql & "GROUP BY T_Etats_Select_Filtre.FiltreName "
            sql = sql & "ORDER BY T_Etats_Select_Filtre.FiltreName; "
            
            
            Set RsFiltre = Con.OpenRecordSet(sql)
            I3 = 0
            If RsFiltre.EOF = True Then
            
            sql = "SELECT T_Etats_Select_Champs.ChampAs, T_indiceProjet.Equipement, T_Etats_Select_Champs.Id_Onglet "
            sql = sql & "FROM T_indiceProjet, T_Etats_Onglet INNER JOIN T_Etats_Select_Champs ON T_Etats_Onglet.Id = T_Etats_Select_Champs.Id_Onglet "
            sql = sql & "WHERE (T_Etats_Select_Champs.ChamsName='OPTION' OR T_Etats_Select_Champs.ChamsName='CRITERES') "
            If Fils = 0 Then
                sql = sql & "AND T_indiceProjet.Id=" & Id_IndiceProjet & " "
            Else
                sql = sql & "AND T_indiceProjet.Id=" & Fils & " "
            End If
            sql = sql & "AND T_Etats_Select_Champs.Id_Onglet=" & Id_Onglet & " AND T_Etats_Onglet.FiltreEquipement=True "
            sql = sql & "GROUP BY T_Etats_Select_Champs.ChamsName, T_Etats_Select_Champs.ChampAs, T_indiceProjet.Equipement, T_indiceProjet.Id, "
            sql = sql & "T_Etats_Select_Champs.Id_Onglet;"
            
            
            Set RsFiltreELab = Con.OpenRecordSet(sql)
            If RsFiltreELab.EOF = False Then
            IsFilterElab = True
            MyOption = FunCodeCritaire(Id_IndiceProjet, "" & RsFiltreELab!Equipement)
                NbMyOption = 0
            '    MyOption = Split("" & RsFiltreELab!Equipement & ";", ";")
            
                Set RangeCopy = SheetFiltre.Range("a1").CurrentRegion
                For I3 = 0 To UBound(MyOption)
                    PoseFiterMyOption = SheetFiltre.Range("a1").CurrentRegion.Rows.Count
                    If Trim("" & MyOption(I3)) <> "" Then
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
                        For I4 = Debut To SheetFiltre.Range("a1").CurrentRegion.Rows.Count
                            For I5 = 1 To SheetFiltre.Range("a1").CurrentRegion.Columns.Count
                                If SheetFiltre.Cells(1, I5) = RsFiltreELab!ChampAs Then
                                    SheetFiltre.Cells(I4, I5) = "*;" & Replace("" & MyOption(I3), "§Null§", "") & ";*"
'                                    SheetFiltre.Cells(I4, I5) = MyOption(I3)
                                    Exit For
                                End If
                            Next
                        Next
                        NbMyOption = NbMyOption + 1
                End If
            Next
            For C = 1 To MyRange.Columns.Count
                If UCase(MyRange(1, I)) = "OPTION" Or UCase(MyRange(1, I)) = UCase("Criteres") Then
                    For L = 2 To MyRange.Rows.Count
                        MyRange(L, C) = ";" & MyRange(L, C) & ";"
                    Next
                    Exit For
                End If
             Next
            Set SheetFiltreCible = IsertSheet(MyExel, MySeet.Name & "_FLT_Equipement")
            FiltreActif MySeet.Range("a1").CurrentRegion, SheetFiltre.Range("A1").CurrentRegion, SheetFiltreCible.Range("a1").CurrentRegion
            
            For C = 1 To MyRange.Columns.Count
                If UCase(MyRange(1, C)) = "OPTION" Or UCase(MyRange(1, C)) = "CRITERES" Then
                    For L = 2 To MyRange.Rows.Count
                        MyRange(L, C) = Left(MyRange(L, C), Len(MyRange(L, C)) - 1)
                        MyRange(L, C) = Right(MyRange(L, C), Len(MyRange(L, C)) - 1)
                    Next
                    Exit For
                End If
            Next
            
            CrerSoutotal SheetFiltreCible, Id_Onglet, Id_IndiceProjet, Fils, NbDoc
            
            
            End If
            End If
            While RsFiltre.EOF = False
            IsFilterElab = True
            Set SheetFiltre = IsertSheet(MyExel, "MyFilterElaborré")
            DeleteRow SheetFiltre, True
            Set MyRange = MySeet.Range("A1").CurrentRegion
            For I3 = 1 To MyRange.Columns.Count
            SheetFiltre.Cells(1, I3) = MyRange(1, I3)
            sql = "SELECT T_Etats_Select_Filtre.FiltreName, T_Etats_Select_Champs.ChampAs, T_Etats_Select_Filtre.Valeur, T_Etats_Select_Filtre.Ligne "
            sql = sql & "FROM (T_Etats_Onglet INNER JOIN T_Etats_Select_Filtre ON T_Etats_Onglet.Id = T_Etats_Select_Filtre.Id_Onglet)  "
            sql = sql & "INNER JOIN T_Etats_Select_Champs ON (T_Etats_Select_Filtre.Colonne = T_Etats_Select_Champs.ChamsName)  "
            sql = sql & "AND (T_Etats_Select_Filtre.Id_Onglet = T_Etats_Select_Champs.Id_Onglet) "
            sql = sql & "WHERE T_Etats_Select_Champs.ChampAs='" & MyReplace(MyRange(1, I3)) & "' "
            sql = sql & "AND T_Etats_Select_Filtre.Id_Onglet=" & Id_Onglet & " AND T_Etats_Select_Filtre.FiltreName='" & MyReplace(RsFiltre!FiltreName) & "' "
            sql = sql & "ORDER BY T_Etats_Select_Filtre.Id;"
            
            Set RsFiltreELab = Con.OpenRecordSet(sql)
            While RsFiltreELab.EOF = False
            
            
            
            SheetFiltre.Cells(RsFiltreELab!Ligne, I3) = "" & RsFiltreELab!Valeur
            SaveNbL = RsFiltreELab!Ligne
            
            
            RsFiltreELab.MoveNext
            Wend
            Next
            
            Set RsFiltreELab = Con.CloseRecordSet(RsFiltreELab)
            For I3 = 2 To SaveNbL
            If SheetFiltre.Range("a1").CurrentRegion.Rows.Count = 1 Then
            SheetFiltre.Rows(CStr(2) & ":" & CStr(2)).Delete Shift:=xlUp
            
            End If
            Next
            sql = "SELECT T_Etats_Select_Champs.ChampAs, T_indiceProjet.Equipement, T_Etats_Select_Champs.Id_Onglet "
            sql = sql & "FROM T_indiceProjet, T_Etats_Onglet INNER JOIN T_Etats_Select_Champs ON T_Etats_Onglet.Id = T_Etats_Select_Champs.Id_Onglet "
            sql = sql & "WHERE (T_Etats_Select_Champs.ChamsName='OPTION' OR T_Etats_Select_Champs.ChamsName='CRITERES') "
           If Fils = 0 Then
                sql = sql & "AND T_indiceProjet.Id=" & Id_IndiceProjet & " "
            Else
                sql = sql & "AND T_indiceProjet.Id=" & Fils & " "
            End If
            sql = sql & "AND T_Etats_Select_Champs.Id_Onglet=" & Id_Onglet & " AND T_Etats_Onglet.FiltreEquipement=True "
            sql = sql & "GROUP BY T_Etats_Select_Champs.ChamsName, T_Etats_Select_Champs.ChampAs, T_indiceProjet.Equipement, T_indiceProjet.Id, "
            sql = sql & "T_Etats_Select_Champs.Id_Onglet;"
            
            
            Set RsFiltreELab = Con.OpenRecordSet(sql)
            If RsFiltreELab.EOF = False Then
            IsFilterElab = True
            MyOption = FunCodeCritaire(Id_IndiceProjet, "" & RsFiltreELab!Equipement)
                NbMyOption = 0
            '    MyOption = Split("" & RsFiltreELab!Equipement & ";", ";")
            
                Set RangeCopy = SheetFiltre.Range("a1").CurrentRegion
                For I3 = 0 To UBound(MyOption)
                    PoseFiterMyOption = SheetFiltre.Range("a1").CurrentRegion.Rows.Count
                    If Trim("" & MyOption(I3)) <> "" Then
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
                        For I4 = Debut To SheetFiltre.Range("a1").CurrentRegion.Rows.Count
                            For I5 = 1 To SheetFiltre.Range("a1").CurrentRegion.Columns.Count
                                If SheetFiltre.Cells(1, I5) = RsFiltreELab!ChampAs Then
                                SheetFiltre.Cells(I4, I5) = "*;" & Replace("" & MyOption(I3), "§Null§", "") & ";*"
'                                    SheetFiltre.Cells(I4, I5) = MyOption(I3)
                                    Exit For
                                End If
                            Next
                        Next
                        NbMyOption = NbMyOption + 1
                End If
            Next
            End If
            For C = 1 To MyRange.Columns.Count
                If UCase(MyRange(1, C)) = "OPTION" Or UCase(MyRange(1, C)) = "CRITERES" Then
                    For L = 2 To MyRange.Rows.Count
                        MyRange(L, C) = ";" & MyRange(L, C) & ";"
                    Next
                    Exit For
                End If
             Next
             For C = 1 To MyRange.Columns.Count
                If UCase(MyRange(1, C)) = "OPTION" Or UCase(MyRange(1, C)) = "CRITERES" Then
                    For L = 2 To MyRange.Rows.Count
                        MyRange(L, C) = ";" & MyRange(L, C) & ";"
                    Next
                    Exit For
                End If
             Next
             Set SheetFiltreCible = IsertSheet(MyExel, MySeet.Name & "_FLT_" & RsFiltre!FiltreName)
            FiltreActif MySeet.Range("a1").CurrentRegion, SheetFiltre.Range("A1").CurrentRegion, SheetFiltreCible.Range("a1").CurrentRegion
            
            
            Set MyRange = SheetFiltreCible.Range("a1").CurrentRegion
            For C = 1 To MyRange.Columns.Count
                If UCase(MyRange(1, C)) = "OPTION" Or UCase(MyRange(1, C)) = "CRITERES" Then
                    For L = 2 To MyRange.Rows.Count
                        MyRange(L, C) = Left(MyRange(L, C), Len(MyRange(L, C)) - 1)
                        MyRange(L, C) = Right(MyRange(L, C), Len(MyRange(L, C)) - 1)
                    Next
                    Exit For
                End If
            Next

            
            CrerSoutotal SheetFiltreCible, Id_Onglet, Id_IndiceProjet, Fils, NbDoc
            
            RsFiltre.MoveNext
            
            Wend
            MySeet.Application.DisplayAlerts = False
            If IsFilterElab = True Then MySeet.Delete
        End If
        
    Next
     
End If

If Trim("" & RsOptionOnglet!OngletStrat) <> "" And RsOptionOnglet!FiltreSequentielle = False Then
    For I = 1 To MyWorkbookSource.Sheets.Count
    Debug.Print UCase(MyWorkbookSource.Sheets(I).Name)
        If UCase(MyWorkbookSource.Sheets(I).Name) = UCase("" & RsOptionOnglet!OngletStrat) Then
            IndexOnglet = IndexOnglet + 1
            Set MySeet = IsertSheet(MyExel, Onglet & "_" & RsOptionOnglet!OngletStrat)
            MyWorkbookSource.Sheets(I).Range("A1").CurrentRegion.Copy
            MyExel.Activate
            MySeet.Select
            MySeet.Range("A1").Select
            MySeet.Paste
            RsSelectChamp.Requery
            Set MyRange = MySeet.Range("A1").CurrentRegion
             RsSelectChamp.Requery
            RsSelectChamp.Filter = "Visible=false"
            While RsSelectChamp.EOF = False
            For I2 = 1 To MyRange.Columns.Count
                If UCase(MyRange(1, I2)) = UCase(RsSelectChamp!ChamsName) Then
                    MyRange(1, I2) = RsSelectChamp!ChampAs
                    MyRange(1, I2).Interior.ColorIndex = 6
                    Exit For
                End If
           Next
            RsSelectChamp.MoveNext
        Wend
        For I2 = MyRange.Columns.Count To 1 Step -1
            If MyRange(1, I2).Interior.ColorIndex <> 6 Then
                DelColonne MySeet, Val(I2)
            End If
         Next
          Set MyRange = MySeet.Range("A1").CurrentRegion
         For I2 = 1 To MyRange.Columns.Count
            If MyRange(1, I2) = "OPTION" Then Exit For
         Next
          For I3 = 2 To MyRange.Rows.Count
            If Trim("" & MyRange(I3, I2)) = "" Then MyRange(I3, I2) = "§Null§"
        Next
         Set MyRange = MySeet.Range("A1").CurrentRegion
          Trier MySeet, NbTrie, Replace("A2:" & MyRange(MyRange.Rows.Count, MyRange.Columns.Count).Address, "$", ""), TrieExcel(1, 0), Val(TrieExcel(1, 1)), TrieExcel(2, 0), Val(TrieExcel(2, 1)), TrieExcel(3, 0), Val(TrieExcel(3, 1))

            Set CollecFiltreName = Nothing
            Set CollecFiltreName = New Collection
            sql = "SELECT T_Etats_Select_Filtre.FiltreName "
            sql = sql & "From T_Etats_Select_Filtre "
            sql = sql & "Where T_Etats_Select_Filtre.Id_Onglet =  " & Id_Onglet & " "
            sql = sql & "GROUP BY T_Etats_Select_Filtre.FiltreName "
            sql = sql & "ORDER BY T_Etats_Select_Filtre.FiltreName; "
            
            
            Set RsFiltre = Con.OpenRecordSet(sql)
            I3 = 0
            If RsFiltre.EOF = True Then
            
            sql = "SELECT T_Etats_Select_Champs.ChampAs, T_indiceProjet.Equipement, T_Etats_Select_Champs.Id_Onglet "
            sql = sql & "FROM T_indiceProjet, T_Etats_Onglet INNER JOIN T_Etats_Select_Champs ON T_Etats_Onglet.Id = T_Etats_Select_Champs.Id_Onglet "
            sql = sql & "WHERE (T_Etats_Select_Champs.ChamsName='OPTION' OR T_Etats_Select_Champs.ChamsName='CRITERES') "
          If Fils = 0 Then
                sql = sql & "AND T_indiceProjet.Id=" & Id_IndiceProjet & " "
            Else
                sql = sql & "AND T_indiceProjet.Id=" & Fils & " "
            End If
            sql = sql & "AND T_Etats_Select_Champs.Id_Onglet=" & Id_Onglet & " AND T_Etats_Onglet.FiltreEquipement=True "
            sql = sql & "GROUP BY T_Etats_Select_Champs.ChamsName, T_Etats_Select_Champs.ChampAs, T_indiceProjet.Equipement, T_indiceProjet.Id, "
            sql = sql & "T_Etats_Select_Champs.Id_Onglet;"
            
            
            Set RsFiltreELab = Con.OpenRecordSet(sql)
            If RsFiltreELab.EOF = False Then
            IsFilterElab = True
            MyOption = FunCodeCritaire(Id_IndiceProjet, "" & RsFiltreELab!Equipement)
                NbMyOption = 0
            '    MyOption = Split("" & RsFiltreELab!Equipement & ";", ";")
            Set SheetFiltre = IsertSheet(MyExel, "MyFilterElaborré")
                Set RangeCopy = SheetFiltre.Range("a1").CurrentRegion
                For I3 = 0 To UBound(MyOption)
                    PoseFiterMyOption = SheetFiltre.Range("a1").CurrentRegion.Rows.Count
                    If Trim("" & MyOption(I3)) <> "" Then
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
                        For I4 = Debut To SheetFiltre.Range("a1").CurrentRegion.Rows.Count
                            For I5 = 1 To SheetFiltre.Range("a1").CurrentRegion.Columns.Count
                                If SheetFiltre.Cells(1, I5) = RsFiltreELab!ChampAs Then
                                SheetFiltre.Cells(I4, I5) = "*;" & Replace("" & MyOption(I3), "§Null§", "") & ";*"
'                                    SheetFiltre.Cells(I4, I5) = MyOption(I3)
                                    Exit For
                                End If
                            Next
                        Next
                        NbMyOption = NbMyOption + 1
                End If
            Next
            For C = 1 To MyRange.Columns.Count
                If UCase(MyRange(1, C)) = "OPTION" Or UCase(MyRange(1, C)) = "CRITERES" Then
                    For L = 2 To MyRange.Rows.Count
                        MyRange(L, C) = ";" & MyRange(L, C) & ";"
                    Next
                    Exit For
                End If
             Next
             Set SheetFiltreCible = IsertSheet(MyExel, MySeet.Name & "_FLT_Equipement")
            FiltreActif MySeet.Range("a1").CurrentRegion, SheetFiltre.Range("A1").CurrentRegion, SheetFiltreCible.Range("a1").CurrentRegion
            
            
            Set MyRange = SheetFiltreCible.Range("a1").CurrentRegion
            For C = 1 To MyRange.Columns.Count
                If UCase(MyRange(1, C)) = "OPTION" Or UCase(MyRange(1, C)) = "CRITERES" Then
                    For L = 2 To MyRange.Rows.Count
                        MyRange(L, C) = Left(MyRange(L, C), Len(MyRange(L, C)) - 1)
                        MyRange(L, C) = Right(MyRange(L, C), Len(MyRange(L, C)) - 1)
                    Next
                    Exit For
                End If
            Next
            CrerSoutotal SheetFiltreCible, Id_Onglet, Id_IndiceProjet, Fils, NbDoc
           
            
            End If
            End If
            While RsFiltre.EOF = False
            IsFilterElab = True
            Set SheetFiltre = IsertSheet(MyExel, "MyFilterElaborré")
            DeleteRow SheetFiltre, True
            Set MyRange = MySeet.Range("A1").CurrentRegion
            For I3 = 1 To MyRange.Columns.Count
            SheetFiltre.Cells(1, I3) = MyRange(1, I3)
            sql = "SELECT T_Etats_Select_Filtre.FiltreName, T_Etats_Select_Champs.ChampAs, T_Etats_Select_Filtre.Valeur, T_Etats_Select_Filtre.Ligne "
            sql = sql & "FROM (T_Etats_Onglet INNER JOIN T_Etats_Select_Filtre ON T_Etats_Onglet.Id = T_Etats_Select_Filtre.Id_Onglet)  "
            sql = sql & "INNER JOIN T_Etats_Select_Champs ON (T_Etats_Select_Filtre.Colonne = T_Etats_Select_Champs.ChamsName)  "
            sql = sql & "AND (T_Etats_Select_Filtre.Id_Onglet = T_Etats_Select_Champs.Id_Onglet) "
            sql = sql & "WHERE T_Etats_Select_Champs.ChampAs='" & MyReplace(MyRange(1, I3)) & "' "
            sql = sql & "AND T_Etats_Select_Filtre.Id_Onglet=" & Id_Onglet & " AND T_Etats_Select_Filtre.FiltreName='" & MyReplace(RsFiltre!FiltreName) & "' "
            sql = sql & "ORDER BY T_Etats_Select_Filtre.Id;"
            
            Set RsFiltreELab = Con.OpenRecordSet(sql)
            While RsFiltreELab.EOF = False
            
            
            
            SheetFiltre.Cells(RsFiltreELab!Ligne, I3) = "" & RsFiltreELab!Valeur
            SaveNbL = RsFiltreELab!Ligne
            
            
            RsFiltreELab.MoveNext
            Wend
            Next
            
            Set RsFiltreELab = Con.CloseRecordSet(RsFiltreELab)
            For I3 = 2 To SaveNbL
            If SheetFiltre.Range("a1").CurrentRegion.Rows.Count = 1 Then
            SheetFiltre.Rows(CStr(2) & ":" & CStr(2)).Delete Shift:=xlUp
            
            End If
            Next
            sql = "SELECT T_Etats_Select_Champs.ChampAs, T_indiceProjet.Equipement, T_Etats_Select_Champs.Id_Onglet "
            sql = sql & "FROM T_indiceProjet, T_Etats_Onglet INNER JOIN T_Etats_Select_Champs ON T_Etats_Onglet.Id = T_Etats_Select_Champs.Id_Onglet "
            sql = sql & "WHERE (T_Etats_Select_Champs.ChamsName='OPTION' OR T_Etats_Select_Champs.ChamsName='CRITERES') "
            If Fils = 0 Then
                sql = sql & "AND T_indiceProjet.Id=" & Id_IndiceProjet & " "
            Else
                sql = sql & "AND T_indiceProjet.Id=" & Fils & " "
            End If
            sql = sql & "AND T_Etats_Select_Champs.Id_Onglet=" & Id_Onglet & " AND T_Etats_Onglet.FiltreEquipement=True "
            sql = sql & "GROUP BY T_Etats_Select_Champs.ChamsName, T_Etats_Select_Champs.ChampAs, T_indiceProjet.Equipement, T_indiceProjet.Id, "
            sql = sql & "T_Etats_Select_Champs.Id_Onglet;"
            
            
            Set RsFiltreELab = Con.OpenRecordSet(sql)
            If RsFiltreELab.EOF = False Then
            IsFilterElab = True
            MyOption = FunCodeCritaire(Id_IndiceProjet, "" & RsFiltreELab!Equipement)
                NbMyOption = 0
            '    MyOption = Split("" & RsFiltreELab!Equipement & ";", ";")
            
                Set RangeCopy = SheetFiltre.Range("a1").CurrentRegion
                For I3 = 0 To UBound(MyOption)
                    PoseFiterMyOption = SheetFiltre.Range("a1").CurrentRegion.Rows.Count
                    If Trim("" & MyOption(I3)) <> "" Then
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
                        For I4 = Debut To SheetFiltre.Range("a1").CurrentRegion.Rows.Count
                            For I5 = 1 To SheetFiltre.Range("a1").CurrentRegion.Columns.Count
                                If SheetFiltre.Cells(1, I5) = RsFiltreELab!ChampAs Then
                                SheetFiltre.Cells(I4, I5) = "*;" & Replace("" & MyOption(I3), "§Null§", "") & ";*"
'                                    SheetFiltre.Cells(I4, I5) = MyOption(I3)
                                    Exit For
                                End If
                            Next
                        Next
                        NbMyOption = NbMyOption + 1
                End If
            Next
            End If
            For C = 1 To MyRange.Columns.Count
                If UCase(MyRange(1, C)) = "OPTION" Or UCase(MyRange(1, C)) = "CRITERES" Then
                    For L = 2 To MyRange.Rows.Count
                        MyRange(L, C) = ";" & MyRange(L, C) & ";"
                    Next
                    Exit For
                End If
             Next
             Set SheetFiltreCible = IsertSheet(MyExel, MySeet.Name & "_FLT_" & RsFiltre!FiltreName)
            FiltreActif MySeet.Range("a1").CurrentRegion, SheetFiltre.Range("A1").CurrentRegion, SheetFiltreCible.Range("a1").CurrentRegion
            
            Set MyRange = SheetFiltreCible.Range("a1").CurrentRegion
            For C = 1 To MyRange.Columns.Count
                If UCase(MyRange(1, C)) = "OPTION" Or UCase(MyRange(1, C)) = "CRITERES" Then
                    For L = 2 To MyRange.Rows.Count
                        MyRange(L, C) = Left(MyRange(L, C), Len(MyRange(L, C)) - 1)
                        MyRange(L, C) = Right(MyRange(L, C), Len(MyRange(L, C)) - 1)
                    Next
                    Exit For
                End If
            Next
            CrerSoutotal SheetFiltreCible, Id_Onglet, Id_IndiceProjet, Fils, NbDoc
                  
            RsFiltre.MoveNext
            
            Wend
            MySeet.Application.DisplayAlerts = False
            If IsFilterElab = True Then
                MySeet.Delete
            Else
                Set MyRange = SheetFiltreCible.Range("a1").CurrentRegion
            For C = 1 To MyRange.Columns.Count
                If UCase(MyRange(1, C)) = "OPTION" Or UCase(MyRange(1, C)) = "CRITERES" Then
                    For L = 2 To MyRange.Rows.Count
                        MyRange(L, C) = Left(MyRange(L, C), Len(MyRange(L, C)) - 1)
                        MyRange(L, C) = Right(MyRange(L, C), Len(MyRange(L, C)) - 1)
                    Next
                    Exit For
                End If
            Next
                CrerSoutotal MySeet, Id_Onglet, Id_IndiceProjet, Fils, NbDoc
            End If
            Exit For
        End If
        
    Next
     
End If
End Function
Function CreatEtaExels(MyExel As EXCEL.Workbook, Rs As Recordset, Onglet As String, Trie As String, _
                        Id_Onglet As Long, Id_IndiceProjet As Long, Fils As Long, Id_Macro As Long, _
                        NbDoc As Long, Optional DapresDoc As Boolean) As Worksheet
Dim MySeet As EXCEL.Worksheet
Dim MyRange As EXCEL.Range
Dim SheetFiltre As Worksheet
Dim SheetFiltreCible As Worksheet
Dim LstOngletName
Dim SaveNbL As Long
Set MySeet = IsertSheet(MyExel, Onglet)
Set MyRange = MySeet.Range("A1").CurrentRegion
Dim RsSelectChamp As Recordset
Dim RsCriteres As Recordset
Dim MyTrie
Dim MyTrieS
Dim sql As String
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
Dim RsChampMasquer As Recordset
Dim KillColImPaire As Boolean
Dim PasFiltre As Boolean
Dim PrefixName As String
PrefixName = "" ' "_FLT_S_Equ_"
ExcelCreatTitre MyRange, Rs
 sql = "SELECT DISTINCT T_Etats_Onglet.Id AS Id_Select, T_Etats_Select_Champs.ChamsName, T_Etats_Select_Champs.Visible "
    sql = sql & "FROM T_Etats_Onglet INNER JOIN T_Etats_Select_Champs ON T_Etats_Onglet.Id = T_Etats_Select_Champs.Id_Onglet "
    sql = sql & "Where T_Etats_Onglet.Id = " & Id_Onglet & " "
    sql = sql & "And T_Etats_Select_Champs.Visible = false; "
    Set RsChampMasquer = Con.OpenRecordSet(sql)

While Rs.EOF = False
    Set MyRange = MySeet.Range("A1").CurrentRegion
    Set MyRange = MySeet.Range(MyRange(MyRange.Rows.Count + 1, 1).Address & ":" & MyRange(MyRange.Rows.Count + 1, MyRange.Columns.Count).Address)
    ExcelCreatTitre MyRange, Rs, True, Formule:=True
    Rs.MoveNext
 Wend
 Set ClsDog = New GenerateurDoc
 ClsDog.RenseigneColonn MySeet, Id_Onglet, 1
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
 sql = "SELECT T_Etats_Select_Filtre.FiltreName "
sql = sql & "From T_Etats_Select_Filtre "
sql = sql & "Where T_Etats_Select_Filtre.Id_Onglet =  " & Id_Onglet & " "
sql = sql & "GROUP BY T_Etats_Select_Filtre.FiltreName "
sql = sql & "ORDER BY T_Etats_Select_Filtre.FiltreName; "

Set RsFiltre = Con.OpenRecordSet(sql)
I = 0
If RsFiltre.EOF = True Then

sql = "SELECT T_Etats_Select_Champs.ChampAs, T_indiceProjet.Equipement, T_Etats_Select_Champs.Id_Onglet "
        sql = sql & "FROM T_indiceProjet, T_Etats_Onglet INNER JOIN T_Etats_Select_Champs ON T_Etats_Onglet.Id = T_Etats_Select_Champs.Id_Onglet "
        sql = sql & "WHERE (T_Etats_Select_Champs.ChamsName='OPTION' OR T_Etats_Select_Champs.ChamsName='CRITERES') "
        If Fils = 0 Then
         sql = sql & "AND T_indiceProjet.Id=" & Id_IndiceProjet & " "
        Else
         sql = sql & "AND T_indiceProjet.Id=" & Fils & " "
        End If
        sql = sql & "AND T_Etats_Select_Champs.Id_Onglet=" & Id_Onglet & " AND T_Etats_Onglet.FiltreEquipement=True "
        sql = sql & "GROUP BY T_Etats_Select_Champs.ChamsName, T_Etats_Select_Champs.ChampAs, T_indiceProjet.Equipement, T_indiceProjet.Id, "
        sql = sql & "T_Etats_Select_Champs.Id_Onglet;"


Set RsFiltreELab = Con.OpenRecordSet(sql)
If RsFiltreELab.EOF = False Then
IsFilterElab = True
MyOption = FunCodeCritaire(Id_IndiceProjet, "" & RsFiltreELab!Equipement)
 Set SheetFiltre = IsertSheet(MyExel, "MyFilterElaborré")
    DeleteRow SheetFiltre, True
    ExcelCreatTitre SheetFiltre.Range("a1").CurrentRegion, Rs
   sql = "SELECT T_Etats_Select_Champs.ChampAs, T_Etats_Select_Champs.CreatOnglet "
    sql = sql & "FROM T_Etats_Onglet INNER JOIN T_Etats_Select_Champs ON T_Etats_Onglet.Id = T_Etats_Select_Champs.Id_Onglet  "
    sql = sql & "Where T_Etats_Select_Champs.Id_Onglet = " & RsFiltreELab!Id_Onglet & "  "
    sql = sql & " And T_Etats_Select_Champs.CreatOnglet=True "
    sql = sql & "GROUP BY T_Etats_Select_Champs.ChampAs, T_Etats_Select_Champs.CreatOnglet, T_Etats_Select_Champs.ChamsName;"

            For I2 = 0 To UBound(MyOption)
            For I3 = 1 To SheetFiltre.Range("a1").CurrentRegion.Columns.Count
                If SheetFiltre.Cells(1, I3) = RsFiltreELab!ChampAs Then
                PrefixName = "_FLT_S_Equ_"
                     SheetFiltre.Cells(I2 + 2, I3) = "*;" & UCase(Replace("" & MyOption(I2), "§Null§", "") & ";*")
    '                SheetFiltre.Cells(I2 + 2, I3) = MyOption(I2)


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

    If Trim(LstOngletName) <> "" Then
    Set SheetFiltre = IsertSheet(MyExel, "MyFilterElaborré")
            If SheetFiltre.Range("A1").CurrentRegion.Rows.Count = 1 And SheetFiltre.Range("A1").CurrentRegion.Columns.Count = 1 Then
                If Trim("" & SheetFiltre.Range("A1")) = "" Then
                    PrefixName = "_Ongt_"
                    For C = 1 To MySeet.Range("a1").CurrentRegion.Columns.Count
                        SheetFiltre.Cells(1, C) = MySeet.Cells(1, C)

                    Next
                End If
             End If
       Set SheetFiltreCible = IsertSheet(MyExel, MySeet.Name & PrefixName, True)
     FiltreActif MySeet.Range("a1").CurrentRegion, SheetFiltre.Range("A1").CurrentRegion, SheetFiltreCible.Range("a1").CurrentRegion
     RsChampMasquer.Requery
     While RsChampMasquer.EOF = False
          DelColonne SheetFiltreCible, ClsDog.MyClonne("" & RsChampMasquer!ChamsName)
        RsChampMasquer.MoveNext
     Wend
        If CrerSoutotal(SheetFiltreCible, Id_Onglet, Id_IndiceProjet, Fils, NbDoc) = False Then
            DeletSheet SheetFiltreCible
        End If
    Else
    
    SplitLstOnglet = Split(ClsDog.LstOngletName & ";", ";")
    If ClsDog.LstOngletName <> "" Then
    For I = 2 To MySeet.Range("a1").CurrentRegion.Rows.Count
        MySeet.Cells(I, ClsDog.MyClonne(ClsDog.LstOngletName)) = ";" & MySeet.Cells(I, ClsDog.MyClonne(ClsDog.LstOngletName)) & ";"
    Next
    End If
            For I4 = 0 To UBound(SplitLstOnglet)
            If Trim(SplitLstOnglet(I4)) <> "" Then
            Set SheetFiltre = IsertSheet(MyExel, "MyFilterElaborré")
            If SheetFiltre.Range("A1").CurrentRegion.Rows.Count = 1 And SheetFiltre.Range("A1").CurrentRegion.Columns.Count = 1 Then
                If Trim("" & SheetFiltre.Range("A1")) = "" Then
                    PrefixName = "_Ongt_"
                    For C = 1 To MySeet.Range("a1").CurrentRegion.Columns.Count
                        SheetFiltre.Cells(1, C) = MySeet.Cells(1, C)
                        SheetFiltre.Cells(2, ClsDog.MyClonne(ClsDog.LstOngletName)) = "a"
                    Next
                End If
            End If

                For I = 2 To SheetFiltre.Range("a1").CurrentRegion.Rows.Count
                    SheetFiltre.Cells(I, ClsDog.MyClonne(ClsDog.LstOngletName)) = ";" & SplitLstOnglet(I4) & ";"
                Next

                        Set SheetFiltreCible = IsertSheet(MyExel, MySeet.Name & PrefixName & SplitLstOnglet(I4), True)


                FiltreActif MySeet.Range("a1").CurrentRegion, SheetFiltre.Range("A1").CurrentRegion, SheetFiltreCible.Range("a1").CurrentRegion

                For I = 2 To SheetFiltreCible.Range("a1").CurrentRegion.Rows.Count
                    SheetFiltreCible.Cells(I, ClsDog.MyClonne(ClsDog.LstOngletName)) = Replace(SheetFiltreCible.Cells(I, ClsDog.MyClonne(ClsDog.LstOngletName)), ";", "")
                Next
                If BoolGenEtatEpisure = True Then

                        VueArriere SheetFiltreCible

                End If
                    RsChampMasquer.Requery
                    KillColImPaire = False
                    SheetFiltreCible.Cells.Replace ";", ""
                While RsChampMasquer.EOF = False
'                SheetFiltreCible.Range("A1").Interior.ColorIndex

                    SheetFiltreCible.Range(Replace(SheetFiltreCible.Cells(1, ClsDog.MyClonne("" & RsChampMasquer!ChamsName)).Address, "$", "")).Interior.ColorIndex = 6
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
'Debute
  While RsFiltre.EOF = False
  IsFilterElab = True
    Set SheetFiltre = IsertSheet(MyExel, "MyFilterElaborré")
    DeleteRow SheetFiltre, True
    ExcelCreatTitre SheetFiltre.Range("a1").CurrentRegion, Rs
    Set MyRange = SheetFiltre.Range("A1").CurrentRegion
 For I = 1 To MyRange.Columns.Count



         sql = "SELECT T_Etats_Select_Filtre.FiltreName, T_Etats_Select_Filtre.Colonne, T_Etats_Select_Filtre.Valeur, T_Etats_Select_Filtre.Ligne "
        sql = sql & "FROM T_Etats_Onglet LEFT JOIN T_Etats_Select_Filtre ON T_Etats_Onglet.Id = T_Etats_Select_Filtre.Id_Onglet "
        sql = sql & "WHERE T_Etats_Select_Filtre.FiltreName='" & MyReplace(RsFiltre!FiltreName) & "' "
        sql = sql & "AND T_Etats_Select_Filtre.Colonne='" & MyReplace(MyRange(1, I)) & "' "
        sql = sql & "AND T_Etats_Select_Filtre.Id_Onglet=" & Id_Onglet & " "
        sql = sql & "ORDER BY T_Etats_Select_Filtre.Id;"




 Set RsFiltreELab = Con.OpenRecordSet(sql)
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
sql = "SELECT T_Etats_Select_Champs.ChampAs, T_indiceProjet.Equipement, T_Etats_Select_Champs.Id_Onglet "
        sql = sql & "FROM T_indiceProjet, T_Etats_Onglet INNER JOIN T_Etats_Select_Champs ON T_Etats_Onglet.Id = T_Etats_Select_Champs.Id_Onglet "
        sql = sql & "WHERE (T_Etats_Select_Champs.ChamsName='OPTION' OR T_Etats_Select_Champs.ChamsName='CRITERES') "
        If Fils = 0 Then
         sql = sql & "AND T_indiceProjet.Id=" & Id_IndiceProjet & " "
         Else
            sql = sql & "AND T_indiceProjet.Id=" & Fils & " "
        End If
        sql = sql & "AND T_Etats_Select_Champs.Id_Onglet=" & Id_Onglet & " AND T_Etats_Onglet.FiltreEquipement=True "
        sql = sql & "GROUP BY T_Etats_Select_Champs.ChamsName, T_Etats_Select_Champs.ChampAs, T_indiceProjet.Equipement, T_indiceProjet.Id, "
        sql = sql & "T_Etats_Select_Champs.Id_Onglet;"


Set RsFiltreELab = Con.OpenRecordSet(sql)
If RsFiltreELab.EOF = False Then

MyOption = FunCodeCritaire(Id_IndiceProjet, "" & RsFiltreELab!Equipement)
    NbMyOption = 0
'    MyOption = Split("" & RsFiltreELab!Equipement & ";", ";")

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
If Trim(ClsDog.LstOngletName) = "" Then
     Set SheetFiltreCible = IsertSheet(MyExel, MySeet.Name & "_FLT_" & RsFiltre!FiltreName, True)
    FiltreActif MySeet.Range("a1").CurrentRegion, SheetFiltre.Range("A1").CurrentRegion, SheetFiltreCible.Range("a1").CurrentRegion

    Set MyRange = SheetFiltreCible.Range("A1").CurrentRegion
        If BoolGenEtatEpisure = True Then
                    VueArriere SheetFiltreCible
                End If
                SheetFiltreCible.Cells.Replace ";", ""
                    RsChampMasquer.Requery
                    KillColImPaire = False
                While RsChampMasquer.EOF = False
'                SheetFiltreCible.Range("A1").Interior.ColorIndex
                    SheetFiltreCible.Range(Replace(SheetFiltreCible.Cells(1, ClsDog.MyClonne("" & RsChampMasquer!ChamsName)).Address, "$", "")).Interior.ColorIndex = 6
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
'FinRd
Else
    SplitLstOnglet = Split(ClsDog.LstOngletName & ";", ";")
    For I = 2 To MySeet.Range("a1").CurrentRegion.Rows.Count
        MySeet.Cells(I, ClsDog.MyClonne(ClsDog.LstOngletName)) = ";" & MySeet.Cells(I, ClsDog.MyClonne(ClsDog.LstOngletName)) & ";"

    Next

    For I4 = 0 To UBound(SplitLstOnglet)
        If Trim(SplitLstOnglet(I4)) <> "" Then
            For I = 2 To SheetFiltre.Range("a1").CurrentRegion.Rows.Count
                SheetFiltre.Cells(I, ClsDog.MyClonne(ClsDog.LstOngletName)) = ";" & SplitLstOnglet(I4) & ";"
            Next
            Set SheetFiltreCible = IsertSheet(MyExel, MySeet.Name & "_FLT_" & RsFiltre!FiltreName & "_" & SplitLstOnglet(I4), True)
            FiltreActif MySeet.Range("a1").CurrentRegion, SheetFiltre.Range("A1").CurrentRegion, SheetFiltreCible.Range("a1").CurrentRegion
             For I = 2 To SheetFiltreCible.Range("a1").CurrentRegion.Rows.Count
                    SheetFiltreCible.Cells(I, ClsDog.MyClonne(ClsDog.LstOngletName)) = Replace(SheetFiltreCible.Cells(I, ClsDog.MyClonne(ClsDog.LstOngletName)), ";", "")
                Next
                If BoolGenEtatEpisure = True Then
                    VueArriere SheetFiltreCible
                End If
                    RsChampMasquer.Requery
                    KillColImPaire = False
                While RsChampMasquer.EOF = False
'                SheetFiltreCible.Range("A1").Interior.ColorIndex
                    SheetFiltreCible.Range(Replace(SheetFiltreCible.Cells(1, ClsDog.MyClonne("" & RsChampMasquer!ChamsName)).Address, "$", "")).Interior.ColorIndex = 6
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
            SheetFiltreCible.Cells(I, ClsDog.MyClonne(ClsDog.LstOngletName)) = Replace(SheetFiltreCible.Cells(I, ClsDog.MyClonne(ClsDog.LstOngletName)), ";", "")
            Next





If CrerSoutotal(SheetFiltreCible, Id_Onglet, Id_IndiceProjet, Fils, NbDoc, False) = False Then
    DeletSheet SheetFiltreCible
End If


        End If
    Next

End If

'For I = 1 To Myrange.Columns.Count
'    If UCase(Myrange(1, I)) = "OPTION" Or UCase(Myrange(1, I)) = UCase("Criteres") Then
'        For I2 = 2 To Myrange.Rows.Count
'            Myrange(I2, I) = Left(Myrange(I2, I), Len(Myrange(I2, I)) - 1)
'            Myrange(I2, I) = Right(Myrange(I2, I), Len(Myrange(I2, I)) - 1)
'        Next
'        Exit For
'    End If
'Next





RsFiltre.MoveNext

Wend
'MySeet.Application.DisplayAlerts = False
If IsFilterElab = True Then
    MySeet.Delete
Else
    CrerSoutotal MySeet, Id_Onglet, Id_IndiceProjet, Fils, NbDoc
End If
Set CreatEtaExels = MySeet
End Function
Function FunCodeCritaire(Id_IndiceProjet As Long, Critere As String)
  Dim Criteres
  Dim Criteres2
  Dim RsCriteres As Recordset
  Dim sql As String
  Dim OpTionValue
  Dim T() As String
  Dim T2() As String
  Dim I As Long
  Dim I2 As Long
  Dim NbCritere As Long
  NbCritere = 0
  ReDim T(NbCritere)
  Criteres = Split("" & Critere & ";", ";")
    For I = 0 To UBound(Criteres)
        If Trim("" & Criteres(I)) <> "" Then
            OpTionValue = Split(Criteres(I) & "_", "_")
                sql = "SELECT T_Critères.CODE_CRITERE, T_Critères.CRITERES From T_Critères "
                sql = sql & "WHERE T_Critères.Id_IndiceProjet=" & Id_IndiceProjet & " "
                sql = sql & "AND (';' & [T_Critères].[CRITERES] &';'  Like '%;" & MyReplace("" & OpTionValue(0)) & ";%' or ';' & [T_Critères].[CODE_CRITERE] & ';' Like '%;" & MyReplace("" & OpTionValue(0)) & ";%') AND T_Critères.ACTIVER=True;"
                 Set RsCriteres = Con.OpenRecordSet(sql)
                While RsCriteres.EOF = False
                   ReDim Preserve T(NbCritere)
                     T(NbCritere) = "" & RsCriteres!CODE_CRITERE
                   NbCritere = NbCritere + 1
                   ReDim Preserve T(NbCritere)
                     T(NbCritere) = "" & RsCriteres!Criteres
                   NbCritere = NbCritere + 1
                   RsCriteres.MoveNext
                Wend
                Set RsCriteres = Con.CloseRecordSet(RsCriteres)
        End If
   Next
   T2 = T
   For I = 0 To UBound(T)
        For I2 = 0 To UBound(T)
            If I <> I2 Then
                
                 ReDim Preserve T2(NbCritere)
                 T2(NbCritere) = T(I) & ";" & T(I2)
                  NbCritere = NbCritere + 1
            End If
        Next
   Next
   ReDim Preserve T2(NbCritere + 2)
                 T2(NbCritere) = "Tous"
                 T2(NbCritere + 1) = "ALL"
                 T2(NbCritere + 2) = "§Null§"
FunCodeCritaire = T2
End Function
