Attribute VB_Name = "ModuleFoctionExcel"
Public PasEncadrer As Boolean
Public Function RechercheParLigne()
  Rows("1:1").Select
    Selection.Find(What:="q0", After:=ActiveCell, LookIn:=xlValues, LookAt _
        :=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:= _
        False).Activate
End Function
Public Sub MisEnFormeXls(MyRangeEncadrer As String, MyExcel As Excel.Application, MyColloneAutoFit As String, MyColloneRepiquer As String, CelluleDepart As String, CelluleTitre As String, CelluleVolet As String, MySheet As Excel.Worksheet, MyLeftHeader As String, _
                         MyCenterHeader As String, MyRightHeader As Date, _
                         MyLeftFooter As String, MyCenterFooter As String, MyRightFooter As String, _
                         MyOrientation As Integer, MyAutoFitRows As Boolean, MyAutoFit As Boolean, BoolTitre As Boolean, _
                         BoolPrintArea As Boolean, FitToPages As Integer, CelluleDepartImpression As String, _
                         CelluleFintImpression As String, OfsetPrintArea As Long, _
                         CouleurTitre As Long)
                         
'Permet la mise en forme d'une feuille Excel.

Dim myrange As Excel.Range
Dim MyAddress As String
Dim I As Long
Dim LondueurPage As Double
MyExcel.DisplayAlerts = False
'MySheet.Select
'Sélectionne les zones contiguës à la cellule donnée en référence.
Set myrange = MySheet.Range(CelluleDepart).CurrentRegion

'Sauvegarde l'adresse du bandeau de titre de colonne.
' Mise en forme du bandeau.
    If BoolTitre = True Then
        MyAddress = CelluleTitre & ":" & MySheet.Cells(MySheet.Range(CelluleTitre).Row, myrange.Columns.Count).Address
        MySheet.Range(MyAddress).Font.Bold = False
        MySheet.Range(MyAddress).Interior.ColorIndex = CouleurTitre
        MySheet.Range(MyAddress).HorizontalAlignment = xlCenter
    End If
' Mise en forme du tableau.
If MyAutoFit = True Then myrange.EntireColumn.AutoFit
If MyColloneAutoFit <> "" Then MySheet.Columns(MyColloneAutoFit).EntireColumn.AutoFit
If PasEncadrer = False Then
    myrange.Borders(xlDiagonalDown).LineStyle = xlNone
    myrange.Borders(xlDiagonalUp).LineStyle = xlNone
    myrange.Borders(xlEdgeLeft).LineStyle = xlContinuous
    myrange.Borders(xlEdgeTop).LineStyle = xlContinuous
    myrange.Borders(xlEdgeBottom).LineStyle = xlContinuous
    myrange.Borders(xlEdgeRight).LineStyle = xlContinuous
    myrange.Borders(xlInsideVertical).LineStyle = xlContinuous
    myrange.Borders(xlInsideHorizontal).LineStyle = xlContinuous
End If
If MyAutoFitRows = True Then myrange.Rows("1:" & CStr(myrange.Rows.Count)).EntireRow.AutoFit
'Initialisation des paramètre d'impression.
If BoolPrintArea = True Then
    If CelluleFintImpression <> "" Then
        MySheet.PageSetup.PrintArea = CelluleDepartImpression & ":" & CelluleFintImpression
    Else
        MySheet.PageSetup.PrintArea = CelluleDepartImpression & ":" & myrange.Cells(myrange.Rows.Count + OfsetPrintArea, myrange.Columns.Count).Address
    End If
End If
    MySheet.PageSetup.PrintTitleRows = MyAddress '"$" & MyRange(1, 1).Row & ":$" & MyRange(1, 1).Row
    If MyColloneRepiquer <> "" Then MySheet.PageSetup.PrintTitleColumns = MyColloneRepiquer
    With MySheet.PageSetup
        .LeftHeader = MyLeftHeader
        .CenterHeader = MyCenterHeader
        .RightHeader = Format(MyRightHeader, "dd/mm/yyyy")
        .LeftFooter = MyLeftFooter
        .CenterFooter = MyCenterFooter '"&P/&N"
        .RightFooter = MyRightFooter
        .LeftMargin = Application.InchesToPoints(0.196850393700787)
        .RightMargin = Application.InchesToPoints(0.196850393700787)
        .TopMargin = Application.InchesToPoints(0.984251968503937)
        .BottomMargin = Application.InchesToPoints(0.984251968503937)
        .HeaderMargin = Application.InchesToPoints(0.511811023622047)
        .FooterMargin = Application.InchesToPoints(0.511811023622047)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
'        .PrintQuality = 300
        .CenterHorizontally = True
        .CenterVertically = False
        .Orientation = MyOrientation 'xlLandscape
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        If FitToPages = 0 Then
            .Zoom = 85
        Else
           .Zoom = False
           .FitToPagesWide = 1
           .FitToPagesTall = 500
        End If
    End With
    
'Figer les volets.
If CelluleVolet <> "" Then
    MySheet.Activate
    MySheet.Range(CelluleVolet).Select
    MyExcel.ActiveWindow.FreezePanes = True
    End If

'Cellules à encadrer.
If MyRangeEncadrer <> "" Then
 MySheet.Range(MyRangeEncadrer).Borders(xlDiagonalDown).LineStyle = xlNone
    MySheet.Range(MyRangeEncadrer).Borders(xlDiagonalUp).LineStyle = xlNone
     MySheet.Range(MyRangeEncadrer).Borders(xlEdgeLeft).LineStyle = xlNone
    MySheet.Range(MyRangeEncadrer).Borders(xlEdgeTop).LineStyle = xlNone
     MySheet.Range(MyRangeEncadrer).Borders(xlEdgeBottom).LineStyle = xlNone
    MySheet.Range(MyRangeEncadrer).Borders(xlEdgeRight).LineStyle = xlNone
     MySheet.Range(MyRangeEncadrer).Borders(xlInsideVertical).LineStyle = xlNone
    MySheet.Range(MyRangeEncadrer).Borders(xlInsideHorizontal).LineStyle = xlNone
     MySheet.Range(MyRangeEncadrer).Borders(xlDiagonalDown).LineStyle = xlNone
    MySheet.Range(MyRangeEncadrer).Borders(xlDiagonalUp).LineStyle = xlNone
    
   MySheet.Range(MyRangeEncadrer).Borders(xlEdgeLeft).LineStyle = xlContinuous
    MySheet.Range(MyRangeEncadrer).Borders(xlEdgeTop).LineStyle = xlContinuous
    MySheet.Range(MyRangeEncadrer).Borders(xlEdgeBottom).LineStyle = xlContinuous
    MySheet.Range(MyRangeEncadrer).Borders(xlEdgeRight).LineStyle = xlContinuous
     End If
End Sub


Public Sub SaveAsHTML(MyPath As String, MyWorbookName As String, MyWorkbook As Excel.Workbook, MyExcel As Excel.Application)
MyExcel.DisplayAlerts = False
Dim Mydir As String

   MyExcel.Workbooks(MyWorkbook.Name).PublishObjects.Add(xlSourcePrintArea, _
        MyPath & MyWorbookName & ".htm", MyWorbookName, "", xlHtmlStatic, _
        "", "").Publish (True)
MyExcel.DisplayAlerts = True
End Sub
'''Permet de sauvegarder au format HTML.
'MyExcel.DisplayAlerts = False
' MyExcel.ActiveWorkbook.SaveAs FileName:= _
'        MyPath & MyWorbookName & ".htm", _
'        FileFormat:=xlHtml, ReadOnlyRecommended:=False, CreateBackup:=False
'MyExcel.DisplayAlerts = True
'End Sub

Public Sub SaveAsXLS(MyPath As String, MyWorbookName As String, MyWorkbook As Excel.Workbook, MyExcel As Excel.Application)
'Permet de sauvegarder au format XLS.
Dim DirSave As String

MyExcel.DisplayAlerts = False
'On Error Resume Next
'Vérifie l'existence du répertoire.
DirSave = Dir(Left(MyPath, Len(MyPath) - 1), vbDirectory)
If DirSave <> "" Then
  MyExcel.Workbooks(MyWorkbook.Name).SaveAs FileName:= _
        MyPath & MyWorbookName, _
        FileFormat:=xlExcel9795, Password:="", WriteResPassword:="", _
        ReadOnlyRecommended:=False, CreateBackup:=False
End If
MyExcel.DisplayAlerts = True
End Sub
Public Function OuvirXls(MyPath As String, MyClasseur As String, MyExcel As Excel.Application) As Workbook
'Permet l 'ouverture d'un fichier Excel.
Dim DirClasseur As String

'Vérifie l'existence du fichier Excel.
DirClasseur = Dir(MyPath & MyClasseur, vbNormal)
Debug.Print MyPath & MyClasseur
If DirClasseur <> "" Then

    MyExcel.Workbooks.Open FileName:= _
       MyPath & MyClasseur
       Set OuvirXls = MyExcel.ActiveWorkbook
End If
End Function
Public Sub InsertLigne(MySheet As Worksheet, MyLinge As String)
'Permet l'insertion d'une ligne.
MySheet.Rows(MyLinge).Insert Shift:=xlDown
End Sub
Public Sub DeleteColumns(MySheets As Excel.Worksheet, Colonne As String)
'Permet la suppression d'une colonne.
MySheets.Columns(Colonne).Delete Shift:=xlToLeft
End Sub
Public Sub DeleteRows(MySheets As Excel.Worksheet, Ligne As String)
'Permet la suppression d'une Ligne.
   MySheets.Rows(Ligne).Delete Shift:=xlUp
End Sub

Public Sub CopyPen(MyExcel As Excel.Application, MyWorkbooksName As String, MySheetsName As String, ColonneSelect As String, ColonnePaste As String)
'Supprime une colonne et recopie la mise en forme d'une autre colonne.
   On Error Resume Next
    MyExcel.Workbooks(MyWorkbooksName).Sheets(MySheetsName).Columns(ColonneSelect).Delete Shift:=xlToLeft
    MyExcel.Workbooks(MyWorkbooksName).Sheets(MySheetsName).Columns(ColonneSelect).Copy
   
    MyExcel.Workbooks(MyWorkbooksName).Sheets(MySheetsName).Columns(ColonnePaste).PasteSpecial Paste:=xlFormats, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Application.CutCopyMode = False
End Sub

Public Function Recherche(myrange As Excel.Range, MyCellule As Long, strRecherche, Mycolonne As Integer) As Long
'Permet de rechercher une valeur dans un tableau Excel.

On Error Resume Next
Recherche = myrange.Find(What:=strRecherche, After:=myrange.Cells(MyCellule, Mycolonne), _
            LookIn:=xlFormulas, LookAt _
        :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False).Row
        
        
        
        
'         Columns("A:A").Select
'    Selection.Find(What:="57", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
'        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
'        False).Activate
        
        
'        Recherche = Myrange.FindNext(What:=strRecherche, After:=Myrange.Cells(MyCellule + 1, 2), _
                    LookIn:=xlFormulas, LookAt:= _
        xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        ).Row
        
        
        
'        Cells.Find(What:="IDF", After:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
'        xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
'        ).Activate
        
    
        
        
    If Err Then Err.Clear
End Function

Public Function RechercheCol(myrange As Excel.Range, MyCellule As Long, strRecherche, Mycolonne As Integer) As Long
'Permet de rechercher une valeur dans un tableau Excel.

On Error Resume Next
RechercheCol = myrange.Find(What:=strRecherche, After:=myrange.Cells(MyCellule, Mycolonne), _
            LookIn:=xlFormulas, LookAt _
        :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False).Column
        
        
        
        
'         Columns("A:A").Select
'    Selection.Find(What:="57", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
'        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
'        False).Activate
        
        
'        Recherche = Myrange.FindNext(What:=strRecherche, After:=Myrange.Cells(MyCellule + 1, 2), _
                    LookIn:=xlFormulas, LookAt:= _
        xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        ).Row
        
        
        
'        Cells.Find(What:="IDF", After:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
'        xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
'        ).Activate
        
        
        
        
    If Err Then
        Err.Clear
        RechercheCol = 0
    End If
End Function

Public Sub EfaceRange(myrange As Range, NumCell As Long, OfsetFin As Long)
'Efface le contenu d'une suite de cellule.
Dim I As Long
For I = 2 To myrange.Rows.Count - OfsetFin
    myrange(I, NumCell) = ""

Next
End Sub
Public Sub MyTri(MyExcel As Excel.Application, NameWorkbooks As String, NameSheet As String, MyCellule As String, SelectCellule As String, boolxlAscending As Boolean)

On Error Resume Next
If boolxlAscending = True Then
    MyExcel.Workbooks(NameWorkbooks).Sheets(NameSheet).Range(SelectCellule).Sort _
    Key1:=MyExcel.Workbooks(NameWorkbooks).Sheets(NameSheet).Range(MyCellule), _
    Order1:=xlAscending, Header:=xlGuess, _
    OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
Else
    MyExcel.Workbooks(NameWorkbooks).Sheets(NameSheet).Range(SelectCellule).Sort _
    Key1:=MyExcel.Workbooks(NameWorkbooks).Sheets(NameSheet).Range(MyCellule), _
    Order1:=xlDescending, Header:=xlGuess, _
    OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
End If
If Err Then
'    MsgBox Err.Description
    Err.Clear
End If
   
End Sub
Public Sub EnleverProtection(MyExcel As Excel.Application, NameWorkbooks As String)

    MyExcel.Workbooks(NameWorkbooks).Unprotect "1234"
End Sub
Public Sub PoserProtection(MyExcel As Excel.Application, NameWorkbooks As String)

    MyExcel.Workbooks(NameWorkbooks).Protect "1234", Structure:=True, Windows:=False
End Sub
Public Sub SupprimerLigne(MyExcel As Excel.Application, _
                            MyWorkbooksName As String, _
                            NbSheets As Long, _
                            CellDepart As String, _
                            CellSource As Long, _
                            CellCible As String, _
                            Ofset As Long, _
                            NumCellValeur As Long, _
                            NbCellule As Long, _
                            Complément1 As String, _
                            Complément2 As String, _
                            boolCouleur As Boolean, _
                            NameSheets As String, _
                            NuSheetsDepart As Integer, _
                            NumColonneCible As Integer)
Dim myrange As Range
Dim MyRange2 As Range
Dim strText As String
Dim pose As Long
Dim Pose_N As Long
Dim PosePointVirgulec1_1 As Integer
Dim PosePointVirgule2 As Integer
Dim Where1 As Integer
Dim Where2 As Integer
Dim NbTrouve As Long
Dim boolEgal As Boolean
Dim DeletOk As Boolean

Dim I As Long
Dim ICell As Long
Dim NbCoupe As Long

I = 2
Set myrange = MyExcel.Workbooks(MyWorkbooksName).Sheets _
(NameSheets).Range(CellDepart).CurrentRegion

Set myrange = MyExcel.Workbooks(MyWorkbooksName).Sheets _
(NameSheets).Range(myrange(1, 1).Address & ":" & myrange(myrange.Rows.Count, NbCellule).Address)
strText = ""
While myrange(I, 1).Value <> ""
strText = ""
    For ICell = NumCellValeur To NumCellValeur + NbCellule
        strText = strText & Trim(myrange(I, ICell))
    Next ICell
    Dim MyInterior As Long
    If strText = "" Then
    NbTrouve = 0
'    MyRange2
    For i2 = NuSheetsDepart To NbSheets
    Pose_N = 1
    pose = 1
'    DeletOk = False
    While pose >= Pose_N
        pose = ModuleFoctionExcel.Recherche(MyExcel.Workbooks(MyWorkbooksName).Sheets(i2).Columns(NumColonneCible), pose, myrange(I, CellSource), 1)
'        Pose = ModuleFoctionExcel.Recherche(EffectifSurToisMois.Columns(1), 1, MyrecordSet!IDA, 1)

        If pose <> 0 Then
        If Pose_N < pose Then
            Pose_N = pose
        End If
        NbTrouve = NbTrouve + 1
        End If
    PosePointVirgule1 = 1
    While PosePointVirgule1 <> 0
'    PosePointVirgule1 = 1
        PosePointVirgule1 = InStr(PosePointVirgule1, Complément1, ";")
        If PosePointVirgule1 <> 0 Then
            Where1 = Val(Mid(Complément1, PosePointVirgule1 + 1, 1))
            Where2 = Val(Mid(Complément2, PosePointVirgule1 + 1, 1))
                If pose <> 0 Then
                    If myrange(I, Where1) = MyExcel.Workbooks(MyWorkbooksName).Sheets(i2).Cells(pose, Where2) Then
                        boolEgal = True
                         MyInterior2 = MyExcel.Workbooks(MyWorkbooksName).Sheets(i2).Cells(pose, 1).Interior.PatternColorIndex
                        MyInterior = MyExcel.Workbooks(MyWorkbooksName).Sheets(i2).Cells(pose, 1).Interior.Pattern
                     Else
                     boolEgal = False
                     
                    End If
                    
                    Else
                    boolEgal = False
                End If
                PosePointVirgule1 = PosePointVirgule1 + 1
          Else
            If Len(Trim(Complément1)) = 0 And pose <> 0 Then
            
                boolEgal = True
            End If
           
        End If
       
    Wend
    If boolEgal = True Then
        ModuleFoctionExcel.DeleteRows MyExcel.Workbooks(MyWorkbooksName).Sheets(i2), pose & ":" & pose
        DeletOk = True
        boolEgal = False
        If boolCouleur = True Then
'            If NbTrouve = 1 Then
'                If Pose > 2 Then
                MyExcel.Workbooks(MyWorkbooksName).Sheets(i2).Rows(pose & ":" & pose).Interior.Pattern = MyInterior
                MyExcel.Workbooks(MyWorkbooksName).Sheets(i2).Rows(pose & ":" & pose).Interior.PatternColorIndex = MyInterior2
'                MyExcel.Workbooks(MyWorkbooksName).Sheets(i2).Rows(Pose + 1 & ":" & Pose + 1).Interior.Pattern = MyInterior
'                Else
'                MyExcel.Workbooks(MyWorkbooksName).Sheets(i2).Rows(Pose + 1 & ":" & Pose + 1).Interior.ColorIndex = 0
'                End If
'            End If
        End If
'
   
    Else
        If pose <= Pose_N Then
        pose = 0
        End If
    End If
    Wend
    
    Next i2
    
  
    ModuleFoctionExcel.DeleteRows MyExcel.Workbooks(MyWorkbooksName).Sheets(NameSheets), I + Ofset & ":" & I + Ofset
    MyExcel.Workbooks(MyWorkbooksName).Sheets(NameSheets).Rows(I + Ofset & ":" & I + Ofset).Interior.Pattern = MyInterior
    MyExcel.Workbooks(MyWorkbooksName).Sheets(NameSheets).Rows(I + Ofset & ":" & I + Ofset).Interior.PatternColorIndex = MyInterior2
    I = I - 1
   
    End If
    I = I + 1
Wend


End Sub

Public Sub MyKill(Repertoir As String, Fichier As String, Extension As String)
On Error Resume Next

Dim strExiste As String

strExiste = Dir(Repertoir & Fichier & Extension, vbNormal)
If strExiste <> "" Then
Kill Repertoir & Fichier & Extension
End If
strExiste = Dir(Repertoir & Fichier & "_fichiers", vbDirectory)

If strExiste <> "" Then
Kill Repertoir & Fichier & "_fichiers\*.*"


RmDir Repertoir & Fichier & "_fichiers"  ' Supprime MONREP."
End If
'INAP0101_fichiers
End Sub
Public Function funMonth(Mois As Integer, An As Integer)
Dim NumMois As Integer
Select Case Mois
Case 1
    funMonth = "'janv-" & Right(CStr(An), 2)
Case 2
    funMonth = "'févr-" & Right(CStr(An), 2)
Case 3
    funMonth = "'mars-" & Right(CStr(An), 2)
Case 4
    funMonth = "'avr-" & Right(CStr(An), 2)
Case 5
    funMonth = "'mai-" & Right(CStr(An), 2)
Case 6
    funMonth = "'juin-" & Right(CStr(An), 2)
Case 7
    funMonth = "'juil-" & Right(CStr(An), 2)
Case 8
    funMonth = "'août-" & Right(CStr(An), 2)
Case 9
    funMonth = "'sept-" & Right(CStr(An), 2)
Case 10
    funMonth = "'Oct-" & Right(CStr(An), 2)
Case 11
    funMonth = "'nov-" & Right(CStr(An), 2)
Case 12
    funMonth = "'déc-" & Right(CStr(An), 2)

End Select

End Function
Public Sub IntiWorkBooks(StrDate As String, MyClasseur As String, MyFeuille As String, MyRep As String, strWorBooksSource As String)
 Dim An As Integer
    Dim Mois As Integer

         Mois = Mid(StrDate, 1, 2)
         An = Right(Trim(StrDate), 4)
      
         BoolTrimestre = False
        If Mois = 2 Or Mois = 4 Or Mois = 6 Or Mois = 8 Or Mois = 10 Or Mois = 12 Then
               BoolTrimestre = True
           End If
        
        MyMonth = funMonth(Mois, An)
        NameWorBooks = MyClasseur & Right(CStr(An), 2) & Format(Mois, "00")
        MyDateClasseur = Right(CStr(An), 2) & Format(Mois, "00")
        NameSheets = MyFeuille & Right(CStr(An), 2) & Format(Mois, "00")
        MyPath = PathCommun & "Vag" & Right(CStr(An), 2) & "-" & Format(Mois, "00") & "\" & MyRep & Right(CStr(An), 2) & "-" & Format(Mois, "00") & "\"
        MyDatePath = Right(CStr(An), 2) & "-" & Format(Mois, "00")
        
       
        WorBooksSource = strWorBooksSource & Right(CStr(An), 2) & Format(Mois, "00") & ".XLS"
         If Mois = 1 Then
            Mois = 12
            An = An - 1
         Else
            Mois = Mois - 1
         End If
         
        NameWorBooksVagMoins1 = MyClasseur & Right(CStr(An), 2) & Format(Mois, "00") & ".XLS"
        MyDateClasseur_N = Right(CStr(An), 2) & Format(Mois, "00")
        MyPath_N = PathCommun & "Vag" & Right(CStr(An), 2) & "-" & Format(Mois, "00") & "\" & MyRep & Right(CStr(An), 2) & "-" & Format(Mois, "00") & "\"
        MyDatePath_N = Right(CStr(An), 2) & "-" & Format(Mois, "00")
       
End Sub
Sub ListeChoix(myrange As Range, MyRangeSource As String) ' As Range)

myrange.NumberFormat = "General"
myrange.Validation.Delete
MyRangeSource = "=" & MyRangeSource
   myrange.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="" & MyRangeSource

End Sub
Public Sub AjusteColonne(myrange As Range)
    myrange.MergeCells = False
    With myrange
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    myrange.Merge
    EnCadre myrange
End Sub
