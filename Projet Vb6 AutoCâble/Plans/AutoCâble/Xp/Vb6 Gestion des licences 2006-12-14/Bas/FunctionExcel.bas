Attribute VB_Name = "FunctionExcel"
Option Explicit
Public Function SerchXlsColumn(MyRange, MyCellule, strRecherche) As Long '
On Error Resume Next
SerchXlsColumn = 0
   SerchXlsColumn = MyRange.Cells.Find(What:=strRecherche, After:=MyCellule, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Column
End Function

Sub DepaceSheet(MyWorkbook As EXCEL.Workbook, NuSheet As Long, NuSheetMov As Long, Optional Fin As Boolean)
    If Fin = True Then
        MyWorkbook.Sheets(NuSheet).Move After:=MyWorkbook.Sheets(NuSheetMov)
    Else
        MyWorkbook.Sheets(NuSheet).Move Before:=MyWorkbook.Sheets(NuSheetMov)
    End If
End Sub
Public Function SuprmerCells(MyRange As Range, Direction As String) As Boolean
On Error Resume Next
SuprmerCells = True
Select Case UCase(Direction)
  
Case "G"
    MyRange.Delete Shift:=xlToLeft
Case "H"
    MyRange.Delete Shift:=xlUp

End Select
If Err Then
Err.Clear
SuprmerCells = False
End If
End Function
Sub ExcelCreatTitre(MyRangeStrate As Object, RsTitre As Recordset, Optional Value As Boolean, Optional ChampEnTire As Boolean, Optional Formule As Boolean)
Dim Row As Long
Dim I As Long
Dim Col As Long
Dim EE
Col = MyRangeStrate.Column
For I = 0 To RsTitre.Fields.Count - 1

    Debug.Print UCase(RsTitre.Fields(I).Name)
'    MyRangeStrate.Application.Visible = True
If Value = False Then
MyRangeStrate(1, Col + I).Select
    MyRangeStrate(1, Col + I) = "'" & UCase(RsTitre.Fields(I).Name)
Else
    If ChampEnTire = True Then
        If RsTitre.Fields(I).Type = 11 Then
            If RsTitre.Fields(I).Value = True Then
                MyRangeStrate(1, 1) = 1
            Else
                MyRangeStrate(1, MyRangeStrate.Columns.Count + 1) = 0
            End If
        Else
            MyRangeStrate(1, 1) = "'" & RsTitre.Fields(I).Value
        End If
    Else
        If RsTitre.Fields(I).Type = 11 Then
            If RsTitre.Fields(I).Value = True Then
                MyRangeStrate(1, Col + I) = 1
            Else
                MyRangeStrate(1, Col + I) = 0
            End If
        Else
            If RsTitre.Fields(I).Type = 5 Or RsTitre.Fields(I).Type = 3 Then
                MyRangeStrate(1, Col + I) = Replace("" & RsTitre.Fields(I).Value, ",", ".")
            Else
                If Formule = True Then
                    EE = Trim("" & RsTitre.Fields(I).Value)
                    If EE <> "" Then
                        If Left(EE, 1) = "=" Then
                            MyRangeStrate(1, Col + I).FormulaR1C1 = RsTitre.Fields(I).Value
                        Else
                            MyRangeStrate(1, Col + I) = "'" & RsTitre.Fields(I).Value
                        End If
                    End If
                Else
                    MyRangeStrate(1, Col + I) = "'" & RsTitre.Fields(I).Value
                End If
            End If
        End If
    End If
 End If
Next
End Sub

Function DeleteRow(MySheet As Worksheet, Optional Tous As Boolean = False)
Dim I As Long
Dim MyRange As Range
Set MyRange = MySheet.Range("A1").CurrentRegion
If Tous = True Then
MySheet.Cells.Delete Shift:=xlUp
Else
For I = MyRange.Rows.Count To 2 Step -1
    MySheet.Rows(CStr(I) & ":" & CStr(I)).Delete Shift:=xlUp
Next
End If
Set MyRange = Nothing
End Function
Function DeletCol(MySheet As Worksheet, Col As String)
MySheet.Columns(Col & ":" & Col).Delete Shift:=xlToLeft
End Function
Public Function DeletPlageCol(MySheet As Worksheet, Col As String)
Dim PointVirgule As Boolean
On Error GoTo ChangeSeparateur
GoTo EttiDelete
ChangeSeparateur:
    Err.Clear
    PointVirgule = True
    If PointVirgule = False Then
        Col = Replace(Col, ",", ";")
    Else
        Col = Replace(Col, ";", ",")
    End If
EttiDelete:
MySheet.Columns(Col).Delete Shift:=xlToLeft
End Function
Sub Trier(MySheet As EXCEL.Worksheet, NbF As Integer, Plage As String, Key1 As String, Order1 As Integer, Key2 As String, Order2 As Integer, Key3 As String, Order3 As Integer)
'
' Trier Macro
' Macro enregistrée le 30/06/2005 par robert.durupt
'
On Error Resume Next
Select Case NbF
    Case 1
        MySheet.Range(Plage).Sort Key1:=MySheet.Range(Key1), Order1:=Order1, Header:= _
        xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortTextAsNumbers
    Case 2
         MySheet.Range(Plage).Sort Key1:=MySheet.Range(Key1), Order1:=Order1, Key2:= _
        MySheet.Range(Key2), Order2:=Order2, Header:=xlGuess, OrderCustom:=1, _
        MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortTextAsNumbers, _
        DataOption2:=xlSortTextAsNumbers
    Case 3
        MySheet.Range(Plage).Sort Key1:=MySheet.Range(Key1), Order1:=Order1, Key2:= _
        MySheet.Range(Key2), Order2:=Order2, Key3:=MySheet.Range(Key3), Order3:=Order3 _
        , Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:= _
        xlTopToBottom, DataOption1:=xlSortTextAsNumbers, DataOption2:=xlSortTextAsNumbers, _
        DataOption3:=xlSortTextAsNumbers
End Select
Err.Clear
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
'MyWorksheet.Application.Visible = True
MyWorksheet.Application.DisplayAlerts = False
            MyWorksheet.Select
            MyWorksheet.Range("A1").CurrentRegion.Replace "§Null§", ""
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
Public Sub ReplaceNull(MySheet As Worksheet, Optional Valeur As String, Optional ValueAs As String)
Dim Value As String
Value = "§Null§"
If Valeur <> "" Then Value = Valeur
'MySheet.Application.Visible = True
MySheet.Select
MySheet.Cells.Replace What:=Value, Replacement:=ValueAs, LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
End Sub
Sub DelColonne(MySheet As Worksheet, I As Long)
Dim MyCol
MyCol = MySheet.Cells(1, I).Columns.Address
MyCol = Split(MyCol, "$")
DeletCol MySheet, "" & MyCol(1)
End Sub
Public Function IsertSheet(MyWorkbook As EXCEL.Workbook, Name As String, Optional Fin As Boolean) As EXCEL.Worksheet
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
Public Function ReplaceBool(MySheet As Worksheet, Plage As String)
On Error Resume Next
Dim NbError As Integer
Dim MyRange As Range
Dim Myrange2 As Range
Dim ColonneName As New Collection
Dim C As Long
Dim C2 As Long
Dim SpliCells
Dim charSplit As String
MySheet.Application.DisplayAlerts = False
charSplit = ","
Set MyRange = MySheet.Range("A1").CurrentRegion

For C = 1 To MyRange.Columns.Count
    ColonneName.Add C, Trim("" & MyRange(1, C))
Next
Reprise:
Set MyRange = MySheet.Range(Plage)
 If NbError > 1 Then
                On Error GoTo 0
                Exit Function
            End If
        If Err Then
            NbError = NbError + 1
            Err.Clear
          Plage = Replace(Plage, ",", ";")
          charSplit = ";"
            GoTo Reprise
            
        End If
       Debug.Print MyRange.Columns.Count
 SpliCells = Split(Plage & charSplit, charSplit)
For C = 0 To UBound(SpliCells) - 1
        Set Myrange2 = MyRange.Range(SpliCells(C))
        For C2 = 1 To Myrange2.Columns.Count
         MyRange.Columns(Myrange2(C2).Column).Replace What:="faux", Replacement:="0", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
       
'
        MyRange.Columns(Myrange2(C2).Column).Replace What:="No", Replacement:="0", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=True, _
        ReplaceFormat:=True
        

          MyRange.Columns(Myrange2(C2).Column).Replace What:="False", Replacement:="0", LookAt:=xlPart, _
          SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
   
         MyRange.Columns(Myrange2(C2).Column).Replace What:="Faux", Replacement:="0", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
        MyRange.Columns(Myrange2(C2).Column).Replace What:="True", Replacement:="1", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        

        MyRange.Columns(Myrange2(C2).Column).Replace What:="VRAI", Replacement:="1", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
         MyRange.Columns(Myrange2(C2).Column).Replace What:="Yes", Replacement:="1", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
     Next
Next
'
'        MySheet.Range(Plage).Replace What:="Yes", Replacement:="1", LookAt:=xlPart, _
'        SearchOrder:=xlByColumns, MatchCase:=False, SearchFormat:=False, _
'        ReplaceFormat:=False
'
'        MySheet.Range(Replace(Myrange(2, ColonneName("TEINT")).Address, "$", "") & ":" & Replace(Myrange(Myrange.Rows.Count, ColonneName("TEINT")).Address, "$", "")).Replace What:="0", Replacement:="NO", LookAt:=xlPart, _
'        SearchOrder:=xlByColumns, MatchCase:=False, SearchFormat:=False, _
'        ReplaceFormat:=False
'
'        MySheet.Range(Replace(Myrange(2, ColonneName("TEINT2")).Address, "$", "") & ":" & Replace(Myrange(Myrange.Rows.Count, ColonneName("TEINT")).Address, "$", "")).Replace What:="0", Replacement:="NO", LookAt:=xlPart, _
'        SearchOrder:=xlByColumns, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
On Error GoTo 0
End Function

