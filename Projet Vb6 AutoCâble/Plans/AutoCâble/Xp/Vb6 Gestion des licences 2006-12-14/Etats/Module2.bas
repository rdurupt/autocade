Attribute VB_Name = "Module2"
Option Explicit

'Function FiltreActif(RangeSource As Range, CriterRange As Range, CopyRange As Range) As Boolean
'FiltreActif = False
'On Error Resume Next
' RangeSource.AdvancedFilter Action:= _
'        xlFilterCopy, CriteriaRange:=CriterRange _
'        , CopyToRange:=CopyRange, Unique:=True
'        DoEvents
'        If Err = 0 Then FiltreActif = True
'        On Error GoTo 0
'End Function

'Function DeleteRow(MySheet As Worksheet, Optional Tous As Boolean = False)
'Dim I As Long
'Dim Myrange As Range
'Set Myrange = MySheet.Range("A1").CurrentRegion
'If Tous = True Then
'MySheet.Cells.Delete Shift:=xlUp
'Else
'For I = Myrange.Rows.Count To 2 Step -1
'    MySheet.Rows(CStr(I) & ":" & CStr(I)).Delete Shift:=xlUp
'Next
'End If
'Set Myrange = Nothing
'End Function

'Sub Trier(MySheet As EXCEL.Worksheet, NbF As Integer, Plage As String, Key1 As String, Order1 As Integer, Key2 As String, Order2 As Integer, Key3 As String, Order3 As Integer)
''
'' Trier Macro
'' Macro enregistrée le 30/06/2005 par robert.durupt
''
'
'Select Case NbF
'    Case 1
'        MySheet.Range(Plage).Sort Key1:=MySheet.Range(Key1), Order1:=Order1, Header:= _
'        xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
'        DataOption1:=xlSortTextAsNumbers
'    Case 2
'         MySheet.Range(Plage).Sort Key1:=MySheet.Range(Key1), Order1:=Order1, Key2:= _
'        MySheet.Range(Key2), Order2:=Order2, Header:=xlGuess, OrderCustom:=1, _
'        MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortTextAsNumbers, _
'        DataOption2:=xlSortTextAsNumbers
'    Case 3
'        MySheet.Range(Plage).Sort Key1:=MySheet.Range(Key1), Order1:=Order1, Key2:= _
'        MySheet.Range(Key2), Order2:=Order2, Key3:=MySheet.Range(Key3), Order3:=Order3 _
'        , Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:= _
'        xlTopToBottom, DataOption1:=xlSortTextAsNumbers, DataOption2:=xlSortTextAsNumbers, _
'        DataOption3:=xlSortTextAsNumbers
'End Select
'
'End Sub

Sub PrintPdf(MyWorbooks As Workbook, Name As String)


'
' Macro3 Macro
' Macro enregistrée le 12/07/2005 par robert.durupt
'

'

    MyWorbooks.Application.ActivePrinter = "PDFCreator sur Ne00:"
'   MyWorbooks.PrintOut From:=1, To:=MyWorbooks.Worksheets.Count, Copies:=1, PrintToFile:=False, _
'         Collate:=True, PrToFileName:=Name
         
  MyWorbooks.PrintOut Copies:=1, PrintToFile:=True, _
          Collate:=True, PrToFileName:=Name
End Sub


