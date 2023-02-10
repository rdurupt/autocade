Attribute VB_Name = "DossierFab"
'Option Explicit

Function DossierDeFab(Xls As String, IdIndiceProjet As Long, Projet As String, Vague As String, Equipement As String, Ensemble As String, PI As String, PL As String, OU As String, LI As String, CLI As String, RefCli As String, F_En_Cours As String, ControlFab As String, NC As String, Croisant As Boolean, Affaire As String, Id_IndiceProjet As Long, Optional SaveAs As String, Optional SaveDoc As Boolean = True) As Boolean
Dim Sql As String
Dim rs As Recordset
Dim I As Long
Dim MyExcel As New EXCEL.Application
MyExcel.DisplayAlerts = False
Dim MyRange As Range
0:  Cell(1) = 0
MyExcel.Visible = True
DoEvents
Set MyWorkbookAppli = MyExcel.Workbooks.Add
'
'  Sql = "SELECT Ligne_Tableau_fils.ACTIVER, Ligne_Tableau_fils.LIAI, Ligne_Tableau_fils.DESIGNATION, "
'  Sql = Sql & "Ligne_Tableau_fils.FIL, Ligne_Tableau_fils.SECT, Ligne_Tableau_fils.TEINT, Ligne_Tableau_fils.TEINT2,  "
'  Sql = Sql & "Ligne_Tableau_fils.ISO, Ligne_Tableau_fils.LONG, Ligne_Tableau_fils.[LONG CP],  "
'  Sql = Sql & "Ligne_Tableau_fils.COUPE, Ligne_Tableau_fils.POS, Ligne_Tableau_fils.[POS-OUT],  "
'  Sql = Sql & "Ligne_Tableau_fils.FA, Ligne_Tableau_fils.APP, Ligne_Tableau_fils.VOI, Ligne_Tableau_fils.PRECO,  "
'  Sql = Sql & "Ligne_Tableau_fils.[Ref Clip], Ligne_Tableau_fils.[Ref Clip Four], Ligne_Tableau_fils.[Ref Joint],  "
'  Sql = Sql & "Ligne_Tableau_fils.[Ref Joint four], Ligne_Tableau_fils.POS2, Ligne_Tableau_fils.[POS-OUT2],  "
'  Sql = Sql & "Ligne_Tableau_fils.FA2, Ligne_Tableau_fils.APP2, Ligne_Tableau_fils.VOI2, Ligne_Tableau_fils.PRECO1,  "
'  Sql = Sql & "Ligne_Tableau_fils.[Ref Clip2], Ligne_Tableau_fils.[Ref Clip Four2], Ligne_Tableau_fils.[Ref Joint2],  "
'  Sql = Sql & "Ligne_Tableau_fils.[Ref Joint Four2], Ligne_Tableau_fils.PRECOG, Ligne_Tableau_fils.OPTION "
'
'   Sql = Sql & "FROM Ligne_Tableau_fils WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & IdIndiceProjet & " ORDER BY val('' & Ligne_Tableau_fils.FIL);"
'
   Sql = "SELECT Ligne_Tableau_fils.ACTIVER, Ligne_Tableau_fils.LIAI, Ligne_Tableau_fils.DESIGNATION,  "
   Sql = Sql & "Ligne_Tableau_fils.FIL, Ligne_Tableau_fils.SECT, Ligne_Tableau_fils.TEINT,  "
   Sql = Sql & "Ligne_Tableau_fils.TEINT2, Ligne_Tableau_fils.ISO, Ligne_Tableau_fils.LONG,  "
   Sql = Sql & "Ligne_Tableau_fils.[LONG CP], Ligne_Tableau_fils.Long_Add,Ligne_Tableau_fils.Long_Add2,Ligne_Tableau_fils.COUPE, Ligne_Tableau_fils.POS,  "
   Sql = Sql & "Ligne_Tableau_fils.[POS-OUT], Ligne_Tableau_fils.FA, Ligne_Tableau_fils.APP,Ligne_Tableau_fils.VOI,Ligne_Tableau_fils.[Ref Connecteur], Ligne_Tableau_fils.[Ref Connecteur_Four],  "
   Sql = Sql & "   "
   Sql = Sql & "Ligne_Tableau_fils.[Ref Clip], Ligne_Tableau_fils.[Ref Clip Four],Ligne_Tableau_fils.PRECO,  "
   Sql = Sql & "Ligne_Tableau_fils.[Ref Joint], Ligne_Tableau_fils.[Ref Joint Four],  "
   Sql = Sql & "Ligne_Tableau_fils.POS2, Ligne_Tableau_fils.[POS-OUT2], Ligne_Tableau_fils.FA2,  "
   Sql = Sql & "Ligne_Tableau_fils.APP2,Ligne_Tableau_fils.VOI2, Ligne_Tableau_fils.[Ref Connecteur2], Ligne_Tableau_fils.[Ref Connecteur_Four2],   "
   Sql = Sql & "Ligne_Tableau_fils.[Ref Clip2], Ligne_Tableau_fils.[Ref Clip Four2],Ligne_Tableau_fils.PRECO2,  "
   Sql = Sql & "Ligne_Tableau_fils.[Ref Joint2], Ligne_Tableau_fils.[Ref Joint Four2],  "
   Sql = Sql & "Ligne_Tableau_fils.PRECOG, Ligne_Tableau_fils.OPTION,  "
   Sql = Sql & "Ligne_Tableau_fils.[Critères spécifiques] "
   Sql = Sql & "FROM Ligne_Tableau_fils "
   Sql = Sql & "Where Ligne_Tableau_fils.Id_IndiceProjet = " & IdIndiceProjet & " "
   Sql = Sql & "ORDER BY Val('' & Ligne_Tableau_fils.FIL);"

Sql = "SELECT Ligne_Tableau_fils.ACTIVER, Ligne_Tableau_fils.LIAI, Ligne_Tableau_fils.DESIGNATION,  "
   Sql = Sql & "Ligne_Tableau_fils.FIL, Ligne_Tableau_fils.SECT, Ligne_Tableau_fils.TEINT,  "
   Sql = Sql & "Ligne_Tableau_fils.TEINT2, Ligne_Tableau_fils.ISO, Ligne_Tableau_fils.LONG,  "
   Sql = Sql & "Ligne_Tableau_fils.[LONG CP], Ligne_Tableau_fils.Long_Add,Ligne_Tableau_fils.Long_Add2,Ligne_Tableau_fils.COUPE, Ligne_Tableau_fils.POS,  "
   Sql = Sql & "Ligne_Tableau_fils.[POS-OUT], Ligne_Tableau_fils.FA, Ligne_Tableau_fils.APP,Ligne_Tableau_fils.VOI,Ligne_Tableau_fils.[Ref Connecteur], Ligne_Tableau_fils.[Ref Connecteur_Four],  "
   Sql = Sql & "   "
   Sql = Sql & "Ligne_Tableau_fils.[Ref Clip], Ligne_Tableau_fils.[Ref Clip Four],Ligne_Tableau_fils.PRECO,  "
   Sql = Sql & "Ligne_Tableau_fils.[Ref Joint], Ligne_Tableau_fils.[Ref Joint Four],  "
   Sql = Sql & "Ligne_Tableau_fils.POS2, Ligne_Tableau_fils.[POS-OUT2], Ligne_Tableau_fils.FA2,  "
   Sql = Sql & "Ligne_Tableau_fils.APP2,Ligne_Tableau_fils.VOI2, Ligne_Tableau_fils.[Ref Connecteur2], Ligne_Tableau_fils.[Ref Connecteur_Four2],   "
   Sql = Sql & "Ligne_Tableau_fils.[Ref Clip2], Ligne_Tableau_fils.[Ref Clip Four2],Ligne_Tableau_fils.PRECO2,  "
   Sql = Sql & "Ligne_Tableau_fils.[Ref Joint2], Ligne_Tableau_fils.[Ref Joint Four2],  "
   Sql = Sql & "Ligne_Tableau_fils.PRECOG, Ligne_Tableau_fils.OPTION,  "
   Sql = Sql & "Ligne_Tableau_fils.[Critères spécifiques] "
   Sql = Sql & "FROM Ligne_Tableau_fils "
   Sql = Sql & "Where Ligne_Tableau_fils.Id_IndiceProjet = " & IdIndiceProjet & " "
   Sql = Sql & "ORDER BY Val('' & Ligne_Tableau_fils.FIL);"
   
   
   
    Set rs = Con.OpenRecordSet(Sql)
    For I = 0 To rs.Fields.Count - 1
    MyWorkbookAppli.Sheets(1).Cells(1, I + 1).Select
        MyWorkbookAppli.Sheets(1).Cells(1, I + 1) = rs.Fields(I).Name
        If rs.Fields(I).Name = "FIL" Then
            Cell(2) = I + 1
        End If
        If rs.Fields(I).Name = "ACTIVER" Then
            MyWorkbookAppli.Sheets(1).Cells(2, I + 1) = 1
            MyWorkbookAppli.Sheets(1).Cells(3, I + 1) = 1
        End If
        If InStr(1, rs.Fields(I).Name, "APP") <> 0 Then
            If Cell(0) = 0 Then
                Cell(0) = I + 1
            Else
                Cell(1) = I + 1
            End If
        End If
    Next
    Set MyRange = MyWorkbookAppli.Sheets(1).Range(MyWorkbookAppli.Sheets(1).Cells(2, Cell(0)).Address & ";" & MyWorkbookAppli.Sheets(1).Cells(3, Cell(1)).Address)
    MyRange.Name = "App"
    ClasseurXls = Xls
    CrateOnglet2 MyExcel, Projet, Affaire, Vague, Equipement, Ensemble, PI, PL, OU, LI, CLI, RefCli, F_En_Cours, ControlFab, NC, Croisant, Id_IndiceProjet, SaveAs, SaveDoc
    
    
    
    
    MyWorkbookAppli.Close False
Set MyWorkbookAppli = Nothing

    MyExcel.Quit
    Set MyExcel = Nothing
    DossierDeFab = True
End Function
Function SheetExiste(MyWorkbook As Workbook, SheetName As String) As Boolean
On Error Resume Next
Dim a
SheetName = Replace(SheetName, "/", "_")
SheetName = Replace(SheetName, "§", "")
a = MyWorkbook.Worksheets(SheetName).Name
If Err = 0 Then SheetExiste = True
DoEvents
End Function
Function FiltreActif(RangeSource As Range, CriterRange As Range, CopyRange As Range) As Boolean
FiltreActif = False
On Error Resume Next
 RangeSource.AdvancedFilter Action:= _
        xlFilterCopy, CriteriaRange:=CriterRange _
        , CopyToRange:=CopyRange, Unique:=True
        DoEvents
        If Err = 0 Then FiltreActif = True
        On Error GoTo 0
End Function
Sub KillRed(MyWorksheet As Worksheet)
Dim MyRange As Range
Dim I
Dim Tous As Long
Set MyRange = MyWorksheet.Range("A1").CurrentRegion
Tous = MyRange.Rows.Count
For I = Tous To 1 Step -1
DoEvents
    If MyRange(I, 4).Font.ColorIndex = 3 Then
        DeleteRows MyWorksheet, CStr(I)
'        I = I - 1
    End If
'    If I > MyRange.Rows.Count Then Exit For
Next
End Sub
Sub DeleteRows(MySheets As EXCEL.Worksheet, Ligne As String)
'Permet la suppression d'une Ligne.
   MySheets.Rows(Ligne).Delete Shift:=xlUp
'   MySheets.Application.Visible = True
End Sub
Sub Kill§(MyWorksheet As Worksheet)
Dim MyRange As Range
Dim I
Set MyRange = MyWorksheet.Range("A1").CurrentRegion

For I = 2 To MyRange.Rows.Count
MyRange(I, Cell(0)) = Replace(MyRange(I, Cell(0)), "§", "")
MyRange(I, Cell(1)) = Replace(MyRange(I, Cell(1)), "§", "")
DoEvents
   
Next
End Sub
Sub ConverRed(MySource As Worksheet, MyCible As Worksheet)
Dim MyRangeSource As Range
Dim MyRangeCible As Range
Dim Pose As Long
Dim I
Set MyRangeSource = MySource.Range("a1").CurrentRegion
Set MyRangeCible = MyCible.Range("a1").CurrentRegion
For I = 2 To MyRangeCible.Rows.Count
Pose = Recherche(MySource.Range(MyRangeSource.Columns(Cell(4)).Address), 1, MyRangeCible(I, Cell(4)).Value, 1)
If Pose <> 0 Then
    MyRangeSource(Pose, Cell(4)).Font.ColorIndex = 3
End If
DoEvents
Next
End Sub
Sub rafraichir_FLT_elaboré(Affaire As String, Piece As String, Liste As String, Ensemble As String, Id_IndiceProjet As Long, Optional Epissur As Boolean = False)
If Trim("" & Replace(MyWorkbookAppli.Worksheets("feuil1").Range("APP").Value, "/", "_")) = "" Then Exit Sub
Dim I
Dim MyPiedTxt
Dim MyRangeSource As Range
 Dim MyWorkbook As Workbook
 Dim RangeTableau As Range
 Dim NewSheet As Worksheet
 
 If SheetExiste(MyWorkbookOnglet, MyWorkbookAppli.Worksheets(1).Range("app").Value) = True Then Exit Sub
 Set NewSheet = MyWorkbookOnglet.Sheets.Add(After:=MyWorkbookOnglet.Worksheets(MyWorkbookOnglet.Worksheets.Count))
 NewSheet.Name = Replace(MyWorkbookAppli.Worksheets(1).Range("app").Value, "/", "_")
 NewSheet.Name = Replace(NewSheet.Name, "§", "")
 
 NewSheet.Select
 Set MyRangeSource = MyWorkbookTravail.Worksheets("Ligne_Tableau_fils").Range("A1").CurrentRegion
' MyRangeSource.Select
 
 FiltreActif MyWorkbookTravail.Worksheets("Ligne_Tableau_fils").Range("A1").CurrentRegion, MyWorkbookAppli.Worksheets(1).Range("A1").CurrentRegion, NewSheet.Cells(1, 1).Range("A1")
'   MyWorkbookTravail.Worksheets("Ligne_Tableau_fils").Range(MyRangeSource.Address).AdvancedFilter Action:= _
'        xlFilterCopy, CriteriaRange:=MyWorkbookAppli.Worksheets("feuil1").Range("A1:V3") _
'        , CopyToRange:=NewSheet.Cells(1, 1).Range("A1"), Unique:=False
        DoEvents
'If UCase(MyWorkbookAppli.Worksheets("feuil1").Range("n2").Value) = "149AA" Then MsgBox ""
If Epissur = False Then KillRed NewSheet
DoEvents
Kill§ NewSheet

   DoEvents
ConverRed MyWorkbookTravail.Worksheets("Ligne_Tableau_fils"), NewSheet
DoEvents
Dim TableFilsD() As String
Dim IndexFilsD As Long
Dim TableFilsG() As String
Dim IndexFilsG As Long
Dim TableFilsC() As String
Dim IndexFilsC As Long
If Left(UCase(NewSheet.Name), 1) = "E" Then Epissur = True
 If Epissur = True Then
    Set MyRangeSource = NewSheet.Cells(1, 1).CurrentRegion
    For I = 2 To MyRangeSource.Rows.Count
      
       If NewSheet.Name = MyRangeSource(I, Cell(0)) Then
            If UCase(Left("" & MyRangeSource(I, Cell(0) + 1) & " ", 1)) = "G" Then
                IndexFilsG = IndexFilsG + 1
                ReDim Preserve TableFilsG(IndexFilsG)
                TableFilsG(IndexFilsG) = MyRangeSource(I, Cell(1)) & " : " & MyRangeSource(I, Cell(1)) & " FILS: " & MyRangeSource(I, Cell(4))
            Else
                 If UCase(Left("" & MyRangeSource(I, Cell(0) + 1) & " ", 1)) = "D" Then
                    IndexFilsD = IndexFilsD + 1
                    ReDim Preserve TableFilsD(IndexFilsD)
                    TableFilsD(IndexFilsD) = MyRangeSource(I, Cell(1)) & " : " & MyRangeSource(I, Cell(1) + 1) & " FILS: " & MyRangeSource(I, Cell(4))
                 Else
                    IndexFilsC = IndexFilsC + 1
                    ReDim Preserve TableFilsC(IndexFilsC)
                    TableFilsC(IndexFilsC) = MyRangeSource(I, Cell(1)) & " : " & MyRangeSource(I, Cell(1) + 1) & " FILS: " & MyRangeSource(I, Cell(4))
                 End If
            End If
       End If
        If NewSheet.Name = MyRangeSource(I, Cell(1)) Then
            If UCase(Left("" & MyRangeSource(I, Cell(1) + 1) & " ", 1)) = "G" Then
                IndexFilsG = IndexFilsG + 1
                ReDim Preserve TableFilsG(IndexFilsG)
                TableFilsG(IndexFilsG) = MyRangeSource(I, Cell(0)) & " : " & MyRangeSource(I, Cell(0) + 1) & " FILS: " & MyRangeSource(I, Cell(4))
            Else
                 If UCase(Left("" & MyRangeSource(I, Cell(1) + 1) & " ", 1)) = "D" Then
                    IndexFilsD = IndexFilsD + 1
                    ReDim Preserve TableFilsD(IndexFilsD)
                    TableFilsD(IndexFilsD) = MyRangeSource(I, Cell(0)) & " : " & MyRangeSource(I, Cell(0) + 1) & " FILS: " & MyRangeSource(I, Cell(4))
                 Else
                    IndexFilsC = IndexFilsC + 1
                    ReDim Preserve TableFilsC(IndexFilsC)
                    TableFilsC(IndexFilsC) = MyRangeSource(I, Cell(0)) & " : " & MyRangeSource(I, Cell(0) + 1) & " FILS: " & MyRangeSource(I, Cell(4))
                 End If
            End If
       End If
    Next
        MyRangeSource(MyRangeSource.Rows.Count + 3, 10) = "Gauche"
        MyRangeSource(MyRangeSource.Rows.Count + 3, 11) = NewSheet.Name
        MyRangeSource(MyRangeSource.Rows.Count + 3, 12) = "Droite"
        FormatExcelPlage MyRangeSource(MyRangeSource.Rows.Count + 3, 10), 40, False, True, xlCenter, xlCenter
        FormatExcelPlage MyRangeSource(MyRangeSource.Rows.Count + 3, 11), 40, False, True, xlCenter, xlCenter
        FormatExcelPlage MyRangeSource(MyRangeSource.Rows.Count + 3, 12), 40, False, True, xlCenter, xlCenter
        For I = 1 To IndexFilsG
        MyRangeSource(MyRangeSource.Rows.Count + 3 + I, 10) = TableFilsG(I)
        Next
        For I = 1 To IndexFilsC
        MyRangeSource(MyRangeSource.Rows.Count + 3 + I, 11) = TableFilsC(I)
        Next
        For I = 1 To IndexFilsD
        MyRangeSource(MyRangeSource.Rows.Count + 3 + I, 12) = TableFilsD(I)
        Next
        I = IndexFilsG
        If I < IndexFilsC Then I = IndexFilsC
        If I < IndexFilsD Then I = IndexFilsD
        
         FormatExcelPlage2 NewSheet.Range(MyRangeSource(MyRangeSource.Rows.Count + 3, 10).Address & ":" & MyRangeSource(MyRangeSource.Rows.Count + 3 + I, 12).Address), 40, False, True, xlCenter, xlCenter, CLng(I)
         If I <> 0 Then I = I + 3
End If
Set MyRangeSource = NewSheet.Cells(1, 1).CurrentRegion


Set RangeTableau = NewSheet.Range("a1").CurrentRegion
          MyPiedTxt = "Debut : __/__/____" & vbCrLf
MyPiedTxt = MyPiedTxt & "Fin : __/__/____" & vbCrLf
MyPiedTxt = MyPiedTxt & "Réalisé par :" & vbCrLf
'Affaire,PI,LI,Ensemble,Equipement,Client

MiseEnPage2 NewSheet, RangeTableau, "Affaire: " & Affaire & Chr(10) & Piece & Chr(10) & Liste, vbCrLf & "Câblage : " & Replace("" & Ensemble, vbCrLf, " ") & Chr(10) & "Equipement : " & Replace("" & Equipement, vbCrLf, " ") & vbCrLf, "Client: " & Client & Chr(10) & Format(Date, "dd-mmm-yyy"), "", "" & MyPiedTxt, "", 2, CLng(I)
'MisEnFormeXls "", Application, "a:y", "", "a1", "", "A2", NewSheet, "", MyWorkbookAppli.Worksheets("feuil1").Range("n2").Value, Date, "", "", "", 2, True, True, False, True, 1, "a1", a(a.Rows.Count, a.Columns.Count).Address, 0, 0
  DoEvents
'If RangeTableau.Rows.Count < 2 Then
'    DeletSheet NewSheet
'End If
Set MyRangeSource = Nothing
Set MyWorkbook = Nothing

Set RangeTableau = Nothing

Set NewSheet = Nothing

End Sub
Sub MiseEnPage2(MyWorksheet As Worksheet, MyRange As Range, MyLeftHeader As String, _
            MyCenterHeader As String, MyRightHeader As String, MyLeftFooter As String, _
            MyCenterFooter As String, MyRightFooter As String, MyxlLandscape As Long, Optional ZonImpressionOfset As Long)
'
' Macro1 Macro
' Macro enregistrée le 27/10/2004 par jerome.ollivon
'
Dim MyZoneImp As Range
Set MyZoneImp = MyWorksheet.Range(MyRange.Cells(1, 1).Address & ":" & MyRange.Cells(MyRange.Rows.Count + ZonImpressionOfset, MyRange.Columns.Count).Address)
MyRange.RowHeight = 200
MyRange.ColumnWidth = 200
MyRange.EntireRow.AutoFit
  MyRange.EntireColumn.AutoFit
  MyRange.Replace "§Null§", ""
If bool_MiseEnPage = True Then
        
        MyWorksheet.PageSetup.PrintArea = Replace(MyZoneImp.Address, "$", "")
        MyWorksheet.PageSetup.LeftHeader = "&""Arial,Normal""&10" & MyLeftHeader
         MyWorksheet.PageSetup.CenterHeader = "&""Arial,Normal""&10" & MyCenterHeader
        MyWorksheet.PageSetup.RightHeader = MyRightHeader
        MyWorksheet.PageSetup.Orientation = xlLandscape
        MyWorksheet.PageSetup.Draft = False
        MyWorksheet.PageSetup.PaperSize = xlPaperA4
        MyWorksheet.PageSetup.FirstPageNumber = xlAutomatic
        MyWorksheet.PageSetup.Order = xlDownThenOver
        MyWorksheet.PageSetup.BlackAndWhite = False
        MyWorksheet.PageSetup.Zoom = False
        MyWorksheet.PageSetup.FitToPagesWide = 1
        MyWorksheet.PageSetup.FitToPagesTall = 1
        MyWorksheet.PageSetup.PrintErrors = xlPrintErrorsDisplayed
        MyWorksheet.PageSetup.TopMargin = MyWorksheet.Application.InchesToPoints(4 / 2.54)
        MyWorksheet.PageSetup.PrintGridlines = False
        MyWorksheet.PageSetup.RightFooter = "&""Arial,Gras""&10&A&""Arial,Normal""&10" & "&10&P/&N"
   
           
           
      DoEvents
         
           
End If
   
End Sub
Sub FormatExcelPlage2(Plage As Range, Couleur As Long, Merge As Boolean, Grille As Boolean, HorizontalAlignment As Long, VerticalAlignment As Long, Optional ZoneImpressionOfset As Long)
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


Sub CherchePrio(MyRangeSource As Range, Valeur As String, Affaire As String, Piece As String, Liste As String, Ensemble As String, Id_IndiceProjet As Long, Optional Epissure As Boolean = False)
Dim Pose As Long
Dim SavePose As Long
Dim MyRange As Range
Dim aa
Pose = 1
SavePose = 0
While Pose <> 0 And Pose > SavePose
SavePose = Pose

Pose = Recherche(MyRangeSource.Range(MyRangeSource.Columns(Cell(0)).Address), Pose, UCase(Valeur), 1, 1)
If Pose <> 0 Then
    If UCase(Left(UCase(Trim("" & MyRangeSource(Pose, Cell(0)).Value)), Len(Valeur))) = UCase(Valeur) Then
    
    On Error Resume Next
    aa = MyRangeSource(Pose, Cell(0)).Value
    aa = 0
    aa = MyCollectionConnecteur("§" & MyRangeSource(Pose, Cell(0)).Value)
    On Error GoTo 0
        If aa <> 0 Then
       
            MyWorkbookAppli.Worksheets(1).Range("app").Value = MyRangeSource(Pose, Cell(0)).Value
            rafraichir_FLT_elaboré Affaire, Piece, Liste, Ensemble, Id_IndiceProjet, Epissure
       End If
    End If
End If
Wend
Pose = 1
SavePose = 0
While Pose <> 0 And Pose > SavePose
SavePose = Pose

Pose = Recherche(MyRangeSource.Range(MyRangeSource.Columns(Cell(1)).Address), Pose, UCase(Valeur), 1, 1)
If Pose <> 0 Then
     If UCase(Left(UCase(Trim("" & MyRangeSource(Pose, Cell(1)).Value)), Len(Valeur))) = UCase(Valeur) Then
         aa = 0
    On Error Resume Next
     aa = MyRangeSource(Pose, Cell(1)).Value
    aa = 0
    aa = MyCollectionConnecteur("§" & MyRangeSource(Pose, Cell(1)).Value)
    On Error GoTo 0
        If aa <> 0 Then
            MyWorkbookAppli.Worksheets(1).Range("app").Value = MyRangeSource(Pose, Cell(1)).Value
            rafraichir_FLT_elaboré Affaire, Piece, Liste, Ensemble, Id_IndiceProjet, Epissure
        End If
    End If
End If
Wend
End Sub
Function OuvirXls(MyClasseur As String, MyExcel As EXCEL.Application) As Workbook
'Permet l 'ouverture d'un fichier Excel.
Dim DirClasseur As String

'Vérifie l'existence du fichier Excel.
DirClasseur = Dir(MyClasseur, vbNormal)
Debug.Print MyClasseur
If DirClasseur <> "" Then

    MyExcel.Workbooks.Open FileName:= _
       MyClasseur
       Set OuvirXls = MyExcel.ActiveWorkbook
End If
End Function
Sub Autofiltre(MyRange As Range)
 MyRange.AutoFilter
End Sub
Sub CrateOnglet2(Application As EXCEL.Application, Projet As String, Affaire As String, Vague As String, Equipement As String, _
                    Ensemble As String, PI As String, PL As String, OU As String, LI As String, CLI As String, RefCli As String, _
                    F_En_Cours As String, ControlFab As String, NC As String, Croisant As Boolean, Id_IndiceProjet As Long, Optional SaveAs As String, Optional SaveDoc As Boolean = True)
Dim Ofset As Long
Dim Sql As String
Dim rs As Recordset
Dim aa
Dim I
Dim FinEpisure As Long
Dim Fso As New FileSystemObject
 'Set MyWorkbookAppli = ActiveWorkbook

If copieclasseur(Application, Projet, Vague, Equipement, Ensemble, PI, PL, OU, LI, CLI, RefCli, F_En_Cours, ControlFab, NC) = False Then Exit Sub
'If Trim("" & Affaire) <> "" Then
'    NumChronoNc = Chrono.Chargement(Affaire)
'        Unload Chrono
'End If
NumChronoNc = FicheNc
 

Dim MyRange As Range

 Set MyCollectionConnecteur = Nothing
Set MyCollectionConnecteur = New Collection
Set MyRange = MyWorkbookTravail.Worksheets("Ligne_Tableau_fils").Range("a1").CurrentRegion

For I = 2 To MyRange.Rows.Count
    MyRange(I, Cell(0)) = "§" & MyRange(I, Cell(0)) & "§"
     MyRange(I, Cell(1)) = "§" & MyRange(I, Cell(1)) & "§"
     MyRange(I, Cell(2)).Font.ColorIndex = 0
Next

Set MyRange = MyWorkbookTravail.Worksheets("Connecteurs").Range("a1").CurrentRegion

For I = 2 To MyRange.Rows.Count
    MyRange(I, Cell(3)) = "§" & MyRange(I, Cell(3)) & "§"
Next

ReDim MyTableau(MyRange.Rows.Count - 1)
Ofset = 1
On Error Resume Next
For I = 2 To MyRange.Rows.Count
    If (UCase(Trim("" & MyRange(I, 2).Value)) = "") Or (UCase(Trim("" & MyRange(I, 1).Value)) = 0) Then
        Ofset = Ofset + 1
    Else
        MyCollectionConnecteur.Add I - Ofset, "§" & MyRange(I, Cell(3)).Value
        MyTableau(MyCollectionConnecteur("§" & MyRange(I, Cell(3)).Value)) = MyRange(I, Cell(3)).Value
        If MyTableau(I - Ofset) = "" Then
            Ofset = Ofset + 1
        End If
    End If
    
Next
On Error GoTo 0
For I = UBound(MyTableau) To LBound(MyTableau) + 1 Step -1
    If Trim("" & MyTableau(I)) = "" Then
        ReDim Preserve MyTableau(I - 1)
    Else
        Exit For
    End If
    
Next
ReDim MyTableaul(UBound(MyTableau), 1)
For I = 0 To UBound(MyTableau)
    MyTableaul(I, 0) = MyTableau(I)
Next

Set MyRange = MyWorkbookTravail.Worksheets("Ligne_Tableau_fils").Range("a1").CurrentRegion
InitBlack MyRange
On Error Resume Next
For I = 2 To MyRange.Rows.Count
    If (UCase(Trim(("" & MyRange(I, 1)))) = 0) Then
        MyRange(I, Cell(2)).Font.ColorIndex = 3
    Else
        aa = 0
        aa = MyCollectionConnecteur("§" & MyRange(I, Cell(0)))
        If aa <> 0 Then
        MyTableaul(MyCollectionConnecteur("§" & MyRange(I, Cell(0))), 1) = CStr(Val(MyTableaul(MyCollectionConnecteur("§" & MyRange(I, Cell(0))), 1)) + 1)
        End If
          aa = 0
        aa = MyCollectionConnecteur("§" & MyRange(I, Cell(1)))
        If aa <> 0 Then
        MyTableaul(MyCollectionConnecteur("§" & MyRange(I, Cell(1))), 1) = CStr(Val(MyTableaul(MyCollectionConnecteur("§" & MyRange(I, Cell(1))), 1)) + 1)
        End If
    End If
Next
On Error GoTo 0
'Set MyRange = ActiveWorkbook.Worksheets("Fil").Range("a5").CurrentRegion
MyTableaul = TriTableau2(MyTableaul)
aa = 0

'§EMH§

'CherchePrio MyRange, ""

CherchePrio MyRange, "§e", Affaire, PI, LI, Ensemble, Id_IndiceProjet, True
CherchePrio MyRange, "'§e", Affaire, PI, LI, Ensemble, Id_IndiceProjet, True
If Croisant = True Then
    DeletSheetEtat MyWorkbookOnglet, 4, 0, 2
Else
    DeletSheetEtat MyWorkbookOnglet, 3, 0, 2
End If
FinEpisure = MyWorkbookOnglet.Sheets.Count + 1
'CherchePrio MyRange, "120"
'CherchePrio MyRange, "'120"
'CherchePrio MyRange, "645"
'CherchePrio MyRange, "'645"
'CherchePrio MyRange, "1337"
'CherchePrio MyRange, "'1337"
'CherchePrio MyRange, "119"
'CherchePrio MyRange, "'119"
'CherchePrio MyRange, "1241"
'CherchePrio MyRange, "'1241"
'CherchePrio MyRange, "1094"
'CherchePrio MyRange, "'1094"
'CherchePrio MyRange, "m"
'CherchePrio MyRange, "'m"
If Croisant = True Then



    For I = 1 To UBound(MyTableaul)
    
        MyWorkbookAppli.Worksheets("Feuil1").Range("APP").Value = MyTableaul(I, 0)
      
        rafraichir_FLT_elaboré Affaire, PI, LI, Ensemble, Id_IndiceProjet
    Next
Else
    For I = UBound(MyTableaul) To 1 Step -1
    
        MyWorkbookAppli.Worksheets("Feuil1").Range("APP").Value = MyTableaul(I, 0)
      
        rafraichir_FLT_elaboré Affaire, PI, LI, Ensemble, Id_IndiceProjet
    Next
End If
Sql = "SELECT RqCartouche.* "
Sql = Sql & "FROM RqCartouche "
 Sql = Sql & "WHERE T_indiceProjet.Id=" & Id_IndiceProjet & ";"
Set rs = Con.OpenRecordSet(Sql)
If rs.EOF = False Then
        PathPl = PathArchive(TableauPath("PathArchiveAutocad"), "" & rs!Client, "" & rs!CleAc, "" & rs!Pieces, "FAB", ControlFab, Val(FormBarGrah.Tag), rs.Fields("PI_Indice"), rs.Fields("LI_Indice"), rs!Version, True)
        PathPl = PathPl & ".XLS"

     MyWorkbookOnglet.Worksheets(1).Select
     If Trim("" & SaveAs) <> "" Then PathPl = SaveAs
     
If Fso.FileExists(PathPl) = True Then Fso.DeleteFile PathPl
Sql = "DELETE T_Dossier_" & EnteteClasseurControle & ".*, T_Dossier_" & EnteteClasseurControle & ".Id_IndiceProjet "
Sql = Sql & "FROM T_Dossier_" & EnteteClasseurControle & " "
Sql = Sql & "WHERE T_Dossier_" & EnteteClasseurControle & ".Id_IndiceProjet=" & Id_IndiceProjet & ";"
Con.Execute Sql


If Croisant = True Then
    DeletSheetEtat MyWorkbookOnglet, 4, 0, FinEpisure
Else
    DeletSheetEtat MyWorkbookOnglet, 3, 0, FinEpisure
End If


DeletSheetEtat MyWorkbookOnglet, 2, 0, 2
ChronoChario = 0
For I = 2 To MyWorkbookOnglet.Sheets.Count
  MyWorkbookOnglet.Sheets(I).Select
'  Set aa = MyWorkbookOnglet.Worksheets(I).Range("A1").CurrentRegion
'  If MyWorkbookOnglet.Worksheets(I).Range("A1").CurrentRegion.Rows.Count = 1 Then
'
'  End If
    insertExelAccess MyWorkbookOnglet.Sheets(I), "T_Dossier_" & EnteteClasseurControle, 1, Id_IndiceProjet, True, True
Next

MyWorkbookOnglet.Sheets(1).Select
MyWorkbookOnglet.SaveAs PathPl, ReadOnlyRecommended:=True

End If
'Dim Fso As New FileSystemObject
Set Fso = Nothing
MyWorkbookOnglet.Close False
MyWorkbookTravail.Close False
Set MyWorkbookTravail = Nothing
Set MyWorkbookOnglet = Nothing
If SaveDoc = True Then _
CreatDoc Affaire, Projet, Ensemble, PI, RefCli, Id_IndiceProjet, F_En_Cours, PL, OU, LI, NC, CLI, SaveDoc

EnteteClasseurControle = EnteteClasseurControle
End Sub
Sub ReplaceWord(Champ As String, Valeur As String)
'
' Macro3 Macro
' Macro enregistrée le 29/10/2004 par robert.durupt
'
  MyWord.Selection.Find.ClearFormatting
    MyWord.Selection.Find.Replacement.ClearFormatting
    With MyWord.Selection.Find
        .Text = Champ
        .Replacement.Text = Valeur
        .Forward = True
        .Wrap = 1
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    MyWord.Selection.Find.Execute
    With MyWord.Selection
        If .Find.Forward = True Then
            .Collapse Direction:=1
        Else
            .Collapse Direction:=0
        End If
        .Find.Execute Replace:=1
        If .Find.Forward = True Then
            .Collapse Direction:=0
        Else
            .Collapse Direction:=1
        End If
        .Find.Execute
    End With
End Sub
Sub ReplaceWordEntete(Champ As String, Valeur As String)

' Macro4 Macro
' Macro enregistrée le 29/10/2004 par robert.durupt
'
    If MyWord.ActiveWindow.View.SplitSpecial <> 0 Then
        MyWord.ActiveWindow.Panes(2).Close
    End If
    If MyWord.ActiveWindow.ActivePane.View.Type = 1 Or MyWord.ActiveWindow. _
        ActivePane.View.Type = 2 Then
        MyWord.ActiveWindow.ActivePane.View.Type = 3
    End If
    MyWord.ActiveWindow.ActivePane.View.SeekView = 9
    MyWord.Selection.Find.ClearFormatting
    MyWord.Selection.Find.Replacement.ClearFormatting
    With MyWord.Selection.Find
        .Text = Champ
        .Replacement.Text = Valeur
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    MyWord.Selection.Find.Execute
    With MyWord.Selection
        If .Find.Forward = True Then
            .Collapse Direction:=1
        Else
            .Collapse Direction:=0
        End If
        .Find.Execute Replace:=1
        If .Find.Forward = True Then
            .Collapse Direction:=0
        Else
            .Collapse Direction:=1
        End If
        .Find.Execute
    End With
    MyWord.ActiveWindow.ActivePane.View.SeekView = 0
End Sub
Sub ReplaceWordEntete2(Champ As String, Valeur As String)
'
' Macro2 Macro
' Macro enregistrée le 18/01/2005 par robert.durupt
'
    If MyWord.ActiveWindow.View.SplitSpecial <> 0 Then
        MyWord.ActiveWindow.Panes(2).Close
    End If
    If MyWord.ActiveWindow.ActivePane.View.Type = 1 Or MyWord.ActiveWindow. _
        ActivePane.View.Type = 2 Then
        MyWord.ActiveWindow.ActivePane.View.Type = 3
    End If
    MyWord.ActiveWindow.ActivePane.View.SeekView = 9
    MyWord.ActiveWindow.ActivePane.View.NextHeaderFooter
    MyWord.Selection.Find.ClearFormatting
    MyWord.Selection.Find.Replacement.ClearFormatting
    With MyWord.Selection.Find
        .Text = Champ
        .Replacement.Text = Valeur
        .Forward = True
        .Wrap = 1
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    MyWord.Selection.Find.Execute
    With MyWord.Selection
        If .Find.Forward = True Then
            .Collapse Direction:=1
        Else
            .Collapse Direction:=0
        End If
        .Find.Execute Replace:=1
        If .Find.Forward = True Then
            .Collapse Direction:=0
        Else
            .Collapse Direction:=1
        End If
        .Find.Execute
    End With
    MyWord.ActiveWindow.ActivePane.View.SeekView = 0
End Sub

Function CreatDoc(Affaire As String, Projet As String, Ensemble As String, Piece As String, PieceCLI As String, Id_IndiceProjet As Long, _
                    F_En_Cours As String, Plan As String, Outil As String, Liste As String, F_NC As String, _
                    Client As String, Optional Save As Boolean = True) As Boolean
Dim Sql As String
Dim rs As Recordset
Dim Fso As New FileSystemObject
CreatDoc = False
Set MyWord = CreateObject("Word.Application")
MyWord.Visible = True
 Set TableauPath = funPath
If Trim("" & TableauPath("Pagegarde")) = "" Then Exit Function
MyWord.Documents.Add DefinirChemienComplet(TableauPath("PathServer"), TableauPath("Pagegarde"))
ReplaceWord """Affaire Num""", Affaire

'ReplaceWord """Affaire Num""", Client
ReplaceWord """ProjetName""", Projet
ReplaceWord """Designation""", Ensemble
ReplaceWord """PI""", Piece
ReplaceWord """NumCli""", PieceCLI

Sql = "SELECT RqCartouche.* "
Sql = Sql & "FROM RqCartouche "
 Sql = Sql & "WHERE T_indiceProjet.Id=" & Id_IndiceProjet & ";"
Set rs = Con.OpenRecordSet(Sql)
If rs.EOF = False Then
        PathPl = PathArchive(TableauPath("PathArchiveAutocad"), "" & rs!Client, "" & rs!CleAc, "" & rs!Pieces, "FAB", "Page_de_Garde", Id_IndiceProjet, rs.Fields("PI_Indice"), rs.Fields("LI_Indice"), rs!Version, True)
        PathPl = PathPl & ".doc"
If Save = True Then
    If Fso.FileExists(PathPl) = True Then Fso.DeleteFile PathPl
    MyWord.ActiveDocument.SaveAs PathPl
End If
End If
'Dim Fso As New FileSystemObject
'PathPl = PathPl & ".XLS"

'aa = PathArchive(TableauPath("PathArchiveAutocad"), FormBarGrah.txt9, FormBarGrah.Affaire, FormBarGrah.txt5, "FAB", "Page_de_Garde", Val(FormBarGrah.Tag), "0", "0", "0", True)

MyWord.ActiveDocument.Close False
'ReplaceWordEntete

'
'
'
'ReplaceWord """Affaire Num""", Outil
If Trim("" & TableauPath("FicheEnCours")) = "" Then Exit Function
MyWord.Documents.Add DefinirChemienComplet(TableauPath("PathServer"), TableauPath("FicheEnCours"))
ReplaceWordEntete """ NumChronoLi """, F_En_Cours
ReplaceWordEntete2 """ NumChronoLi """, F_En_Cours
ReplaceWordEntete """AffaireNum""", Affaire
ReplaceWordEntete2 """AffaireNum""", Affaire
ReplaceWord """Designation""", Ensemble
ReplaceWord """ Ref_Pi_Cli """, PieceCLI
ReplaceWord """PI""", Piece
ReplaceWord """PL""", Plan
ReplaceWord """OU""", Outil
ReplaceWord """LI""", Liste
ReplaceWordEntete """Date""", Date
ReplaceWordEntete2 """Date""", Date

Sql = "SELECT RqCartouche.* "
Sql = Sql & "FROM RqCartouche "
 Sql = Sql & "WHERE T_indiceProjet.Id=" & Id_IndiceProjet & ";"
Set rs = Con.OpenRecordSet(Sql)
If rs.EOF = False Then
        PathPl = PathArchive(TableauPath("PathArchiveAutocad"), "" & rs!Client, "" & rs!CleAc, "" & rs!Pieces, "FAB", F_En_Cours, Id_IndiceProjet, rs.Fields("PI_Indice"), rs.Fields("LI_Indice"), rs!Version, True)
        PathPl = PathPl & ".doc"
If Save = True Then
    If Fso.FileExists(PathPl) = True Then Fso.DeleteFile PathPl

    MyWord.ActiveDocument.SaveAs PathPl
End If
End If


'aa = PathArchive(TableauPath("PathArchiveAutocad"), FormBarGrah.txt9, FormBarGrah.Affaire, FormBarGrah.txt5, "FAB", FormBarGrah.txt13, Val(FormBarGrah.Tag), "0", "0", "0", True)

MyWord.ActiveDocument.Close False

If Trim("" & TableauPath("ModelNC")) = "" Then Exit Function
MyWord.Documents.Add DefinirChemienComplet(TableauPath("PathServer"), TableauPath("ModelNC"))

ReplaceWordEntete """NC""", F_NC
ReplaceWordEntete """Chrono""", F_NC
ReplaceWord """CLIENT""", Client
ReplaceWord """PieceCLI""", PieceCLI
'ReplaceWord """ RefPf """, RefPF
'ReplaceWord """PI""", Piece
'ReplaceWord """OU""", Outil
ReplaceWordEntete """Date""", Date

'ReplaceWord """Affaire Num""", Affaire
FicheNc = Replace(FicheNc, "_Rév.:", "")
If rs.EOF = False Then
    PathPl = PathArchive(TableauPath("PathArchiveAutocad"), "" & rs!Client, "" & rs!CleAc, "" & rs!Pieces, "DNC", F_NC, Id_IndiceProjet, rs.Fields("PI_Indice"), rs.Fields("LI_Indice"), rs!Version, True)
    PathPl = PathPl & ".doc"
If Save = True Then
    If Fso.FileExists(PathPl) = True Then Fso.DeleteFile PathPl

    MyWord.ActiveDocument.SaveAs PathPl
End If
End If
'aa = PathArchive(TableauPath("PathArchiveAutocad"), FormBarGrah.txt9, FormBarGrah.Affaire, FormBarGrah.txt5, "DNC", FormBarGrah.txt16, Val(FormBarGrah.Tag), "0", "0", "0", True)
MyWord.ActiveDocument.Close False
MyWord.Quit
Set MyWord = Nothing
CreatDoc = True
End Function
Function Recherche(MyRange As EXCEL.Range, MyCellule As Long, strRecherche, Mycolonne As Integer, Optional MyxlWhole As Long) As Long
'Permet de rechercher une valeur dans un tableau Excel.
MyxlWhole = MyxlWhole + 1
On Error Resume Next
Recherche = MyRange.Find(What:=strRecherche, After:=MyRange.Cells(MyCellule, Mycolonne), _
            LookIn:=xlFormulas, LookAt _
        :=MyxlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False).Row
        
        
        
 
    If Err Then Err.Clear
End Function
Sub InitBlack(MyRange As Range)
Dim I
For I = 1 To MyRange.Rows.Count
    MyRange(I, Cell(2)).Font.ColorIndex = 0
Next
End Sub
Sub ReplaceShapesTxt(Text As String, ReplacementText As String, MyWord As Object)
'
' Macro6 Macro
' Macro enregistrée le 18/01/2005 par robert.durupt
'
    MyWord.Selection.Find.ClearFormatting
    MyWord.Selection.Find.Replacement.ClearFormatting
    With MyWord.Selection.Find
        .Text = Text
        .Replacement.Text = ReplacementText
        .Forward = True
        .Wrap = 1
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    MyWord.Selection.Find.Execute
    DoEvents
    With MyWord.Selection
        If .Find.Forward = True Then
            .Collapse Direction:=1
        Else
            .Collapse Direction:=0
        End If
        .Find.Execute Replace:=1
        DoEvents
        If .Find.Forward = True Then
            .Collapse Direction:=0
        Else
            .Collapse Direction:=1
        End If
        .Find.Execute
        DoEvents
    End With
  
End Sub
Sub ReplaceShapes(Application As EXCEL.Application, TableauText)
Dim MyWord As Object
Dim MyVerb
Dim I
Set MyVerb = Application.ActiveWorkbook.Worksheets(1).OLEObjects(1)
   MyVerb.Verb Verb:=xlOpen
   'MyVerb.Verb Verb:=xlClosed

   Set MyWord = MyVerb.object.Application.ActiveDocument.Application
    DoEvents
    On Error Resume Next

    On Error GoTo 0
    For I = 1 To UBound(TableauText)
ReplaceShapesTxt "" & TableauText(I, 0), "" & TableauText(I, 1), MyWord.Application
Next
ReplaceShapesTxt """controle""", EnteteClasseurControle, MyWord.Application
MyWord.ActiveWindow.Close
'   Selection.Verb Verb:=xlClosed
If MyWord.Documents.Count = 0 Then
   MyWord.Quit
End If
    Set MyWord = Nothing
'    MyVerb.Verb Verb:=xlClosed

End Sub
Sub PageDeGare(FichierXLS As String, PathAccess As String, Projet As String, Vague As String, Equipement As String, Ensemble As String, PI As String, PL As String, OU As String, LI As String, CLI As String, RefCli As String, F_En_Cours As String, ControlFab As String, NC As String)

If SheetExiste(MyWorkbookOnglet, "page_de_garde") = False Then Exit Sub
Dim MyListe
Dim Sql As String
Dim MyCon As New Ado
Dim rs As Recordset
MyListe = Split(ClasseurXls, "\")
Dim TableauText(2, 1) As String
TableauText(1, 0) = """ NumChrono """: TableauText(1, 1) = ControlFab
TableauText(2, 0) = """date""": TableauText(2, 1) = Format(Date, "dd/mm/yyyy")
 ReplaceShapes MyWorkbookOnglet.Application, TableauText
 
 Ensemble = Replace(Ensemble, Chr(13), "")
 Equipement = Replace(Equipement, Chr(13), "")
 MyWorkbookOnglet.Worksheets("page_de_garde").Range("CLIENT").Value = "CLIENT : " & CLI
MyWorkbookOnglet.Worksheets("page_de_garde").Range("Designation").Value = Ensemble
MyWorkbookOnglet.Worksheets("page_de_garde").Range("Equipement").Value = Equipement
MyWorkbookOnglet.Worksheets("page_de_garde").Range("Ref_PI").Value = "Pièce : " & PI
MyWorkbookOnglet.Worksheets("page_de_garde").Range("Ref_pl").Value = "Plan : " & PL
MyWorkbookOnglet.Worksheets("page_de_garde").Range("Ref_Ou").Value = "Outil : " & OU
MyWorkbookOnglet.Worksheets("page_de_garde").Range("Ref_li").Value = "Liste : " & LI
MyWorkbookOnglet.Worksheets("page_de_garde").Range("Ref_cli").Value = "REF CLI : " & RefCli
Ensemble = Replace(Ensemble, Chr(10), " ")
 Equipement = Replace(Equipement, Chr(10), " ")
End Sub
Function copieclasseur(Application As EXCEL.Application, Projet As String, Vague As String, Equipement As String, Ensemble As String, PI As String, PL As String, OU As String, LI As String, CLI As String, RefCli As String, F_En_Cours As String, ControlFab As String, NC As String) As Boolean
copieclasseur = False
' copieclasseur Macro
' Macro enregistrée le 26/10/2004 par jerome.ollivon
'
Set TableauPath = funPath
Dim MyChdir As String
'
Dim MyRange As Range
'MyChdir = ScanFichier.Chargement("XLS", "")
'Unload ScanFichier
'MyChdir = InputBox("fichier", "entrez votre nom de fichier")
If Trim("" & ClasseurXls) = "" Then Exit Function
   
      Set MyWorkbookTravail = OuvirXls(ClasseurXls, Application)
      MyWorkbookTravail.Activate
      DoEvents
      Set MyRange = MyWorkbookTravail.Sheets("Connecteurs").Range("A1").CurrentRegion
      For I = 1 To MyRange.Columns.Count
      If UCase(MyRange(1, I)) = "CODE_APP" Then
        Cell(3) = I
        
      End If
       If UCase(MyRange(1, I)) = "O/N" Then
        Cell(4) = I
        
      End If
      Next
  Autofiltre MyWorkbookTravail.Worksheets("Ligne_Tableau_fils").Range("A1")

        Set MyWorkbookOnglet = Application.Workbooks.Add(DefinirChemienComplet(TableauPath("PathServer"), TableauPath("ModelControl")))
        MyWorkbookOnglet.Activate
        DoEvents
   DoEvents
    PageDeGare ClasseurXls, Con.RetournDbName("MDB"), Projet, Vague, Equipement, Ensemble, PI, PL, OU, LI, CLI, RefCli, F_En_Cours, ControlFab, NC
    copieclasseur = True
End Function
