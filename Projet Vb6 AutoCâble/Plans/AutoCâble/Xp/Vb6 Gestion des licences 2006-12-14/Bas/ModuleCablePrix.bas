Attribute VB_Name = "ModuleCablePrix"
Sub ImportCablePrix()
Dim MyExcel As New EXCEL.Application
Dim MyClasseur As EXCEL.Workbook
Dim MySheet As EXCEL.Worksheet
Dim MyRange As EXCEL.Range
Dim Sql As String
Dim ErrDesciption As String
Dim Rs As Recordset
FormBarGrah.ProgressBar1Caption.Caption = " Importer Prix Câbles :"
FormBarGrah.ProgressBar1.Value = 0
Set MyExcel = New EXCEL.Application
MyExcel.DisplayAlerts = False
'MyExcel.Visible = True
On Error GoTo Fin
Set MyClasseur = MyExcel.Workbooks.Open(App.Path & "\DossierAplication\ImportPrix\" & XlsPrix & ".xls")
'MyClasseur.Application.Visible = True
Set MyRange = MyClasseur.Worksheets("Prix").Range("A1").CurrentRegion
'Set MyWorSheet = MyWorkbook.Worksheets("")
Sql = "UPDATE " & XlsPrix & " SET " & XlsPrix & ".Supp = True;"
Con.Execute Sql
FormBarGrah.ProgressBar1.Max = MyRange.Rows.Count
For I = 2 To MyRange.Rows.Count
FormBarGrah.ProgressBar1.Value = I
DoEvents
    Sql = "Select " & XlsPrix & ".id "
    Sql = Sql & "From " & XlsPrix & " "
    Sql = Sql & "Where " & XlsPrix & ".Section='" & Replace(MyRange(I, 1), ",", ".") & "' AND " & XlsPrix & ".ISO='" & MyReplace(MyRange(I, 2)) & "';"
    Set Rs = Con.OpenRecordSet(Sql)
    If Rs.EOF = False Then
        Sql = "UPDATE " & XlsPrix & " SET " & XlsPrix & ".[Section] = '" & Replace(MyRange(I, 1), ",", ".") & "', " & XlsPrix & ".ISO = '" & MyReplace(MyRange(I, 2)) & "', "
        Sql = Sql & "" & XlsPrix & ".[Prix U] =" & Replace(MyRange(I, 3), ",", ".") & ", " & XlsPrix & ".Supp = False "
        Sql = Sql & "WHERE " & XlsPrix & ".id=" & Rs!Id & " ;"

    Else
        Sql = "INSERT INTO " & XlsPrix & " ( [Section], ISO, [Prix U] ) "
        Sql = Sql & "VALUES('" & MyRange(I, 1) & "', '" & MyReplace(MyRange(I, 2)) & "'," & Replace(MyRange(I, 3), ",", ".") & ");"

    End If
    Con.Execute Sql
Next
Sql = "DELETE " & XlsPrix & ".*, " & XlsPrix & ".Supp FROM " & XlsPrix & " WHERE " & XlsPrix & ".Supp=True;"
Con.Execute Sql

MyClasseur.Close False
Fin:
ErrDesciption = Err.Description
If Trim("" & ErrDesciption) <> "" Then
    MsgBox ErrDesciption
End If
Set MyRange = Nothing
Set MyClasseur = Nothing
MyExcel.Quit
Set MyExcel = Nothing
FormBarGrah.ProgressBar1Caption.Caption = " Fin du Tratement:"
FormBarGrah.ProgressBar1.Value = 0
DoEvents
End Sub
Sub ExportCablePrix()
Dim Sql As String
Dim Rs As Recordset
Dim I As Long
Dim Fso As New FileSystemObject
FormBarGrah.ProgressBar1Caption.Caption = " Exorter Prix Câbles :"
FormBarGrah.ProgressBar1.Value = 0
Sql = "SELECT " & XlsPrix & ".ISO, " & XlsPrix & ".Section, " & XlsPrix & ".[Prix u] "
Sql = Sql & "FROM " & XlsPrix & " "
Sql = Sql & "ORDER BY " & XlsPrix & ".ISO, " & XlsPrix & ".Section;"
Set Rs = Con.OpenRecordSet(Sql)

Dim MyExcel As New EXCEL.Application
MyExcel.DisplayAlerts = False
Dim MyClasseur As EXCEL.Workbook
Dim MySheet As EXCEL.Worksheet
Dim MyRange As EXCEL.Range

If Fso.FileExists(App.Path & "\DossierAplication\ExportPrix\" & XlsPrix & ".xls") = True Then
Fso.DeleteFile App.Path & "\DossierAplication\ExportPrix\" & XlsPrix & ".xls"
End If

'MyExcel.Visible = True

Set MyClasseur = MyExcel.Workbooks.Add(App.Path & "\DossierAplication\ModèlePrix\ModèlePrix.xlt")

While Rs.EOF = False
I = I + 1
    Rs.MoveNext
Wend
If I = 0 Then I = 1
FormBarGrah.ProgressBar1.Max = I
Rs.Requery
I = 1
While Rs.EOF = False
I = I + 1
    IncremanteBarGrah FormBarGrah
    DoEvents
     MyClasseur.Worksheets("Prix").Cells(I, 1) = Val(Replace("" & Rs!Section, ",", "."))
     MyClasseur.Worksheets("Prix").Cells(I, 2) = "" & Rs!ISO
     MyClasseur.Worksheets("Prix").Cells(I, 3) = Val(Replace("" & Rs![Prix u], ",", "."))
    Rs.MoveNext
Wend
MyClasseur.SaveAs App.Path & "\DossierAplication\ExportPrix\" & XlsPrix & ".xls", ReadOnlyRecommended:=True
FormBarGrah.ProgressBar1Caption.Caption = " Fin du Tratement:"
FormBarGrah.ProgressBar1.Value = 0
DoEvents

MyClasseur.Close False
Set MyClasseur = Nothing
MyExcel.Quit
Set MyExcel = Nothing
End Sub

