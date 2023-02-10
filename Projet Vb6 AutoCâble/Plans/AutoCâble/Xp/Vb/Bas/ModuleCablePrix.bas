Attribute VB_Name = "ModuleCablePrix"
Public Sub ImportCablePrix()
Dim MyEcel As New EXCEL.Application
Dim MyClasseur As EXCEL.Workbook
Dim MySheet As EXCEL.Worksheet
Dim Myrange As EXCEL.Range
Dim Sql As String
Dim ErrDesciption As String
Dim Rs As Recordset
FormBarGrah.ProgressBar1Caption.Caption = " Importer Prix Câbles :"
FormBarGrah.ProgressBar1.Value = 0
Set MyExcel = New EXCEL.Application
'MyEcel.Visible = True
On Error GoTo Fin
Set MyClasseur = MyEcel.Workbooks.Open(App.Path & "\ImportPrix\" & XlsPrix & ".xls")

Set Myrange = MyClasseur.Worksheets("Prix").Range("A1").CurrentRegion
'Set MyWorSheet = MyWorkbook.Worksheets("")
Sql = "UPDATE " & XlsPrix & " SET " & XlsPrix & ".Supp = True;"
Con.Exequte Sql
FormBarGrah.ProgressBar1.Max = Myrange.Rows.Count
For i = 2 To Myrange.Rows.Count
FormBarGrah.ProgressBar1.Value = i
DoEvents
    Sql = "Select " & XlsPrix & ".id "
    Sql = Sql & "From " & XlsPrix & " "
    Sql = Sql & "Where " & XlsPrix & ".Section='" & Myrange(i, 1) & "' AND " & XlsPrix & ".ISO='" & MyReplace(Myrange(i, 2)) & "';"
    Set Rs = Con.OpenRecordSet(Sql)
    If Rs.EOF = False Then
        Sql = "UPDATE " & XlsPrix & " SET " & XlsPrix & ".[Section] = '" & (Myrange(i, 1)) & "', " & XlsPrix & ".ISO = '" & MyReplace(Myrange(i, 2)) & "', "
        Sql = Sql & "" & XlsPrix & ".[Prix U] =" & Replace(Myrange(i, 3), ",", ".") & ", " & XlsPrix & ".Supp = False "
        Sql = Sql & "WHERE " & XlsPrix & ".id=" & Rs!Id & " ;"

    Else
        Sql = "INSERT INTO " & XlsPrix & " ( [Section], ISO, [Prix U] ) "
        Sql = Sql & "VALUES('" & Myrange(i, 1) & "', '" & MyReplace(Myrange(i, 2)) & "'," & Replace(Myrange(i, 3), ",", ".") & ");"

    End If
    Con.Exequte Sql
Next
Sql = "DELETE " & XlsPrix & ".*, " & XlsPrix & ".Supp FROM " & XlsPrix & " WHERE " & XlsPrix & ".Supp=True;"
Con.Exequte Sql

MyClasseur.Close False
Fin:
ErrDesciption = Err.Description
If Trim("" & ErrDesciption) <> "" Then
    MsgBox ErrDesciption
End If
Set Myrange = Nothing
Set MyClasseur = Nothing
MyEcel.Quit
Set MyEcel = Nothing
FormBarGrah.ProgressBar1Caption.Caption = " Fin du Tratement:"
FormBarGrah.ProgressBar1.Value = 0
DoEvents
End Sub
Public Sub ExportCablePrix()
Dim Sql As String
Dim Rs As Recordset
Dim i As Long
Dim Fso As New FileSystemObject
FormBarGrah.ProgressBar1Caption.Caption = " Exorter Prix Câbles :"
FormBarGrah.ProgressBar1.Value = 0
Sql = "SELECT " & XlsPrix & ".ISO, " & XlsPrix & ".Section, " & XlsPrix & ".[Prix u] "
Sql = Sql & "FROM " & XlsPrix & " "
Sql = Sql & "ORDER BY " & XlsPrix & ".ISO, " & XlsPrix & ".Section;"
Set Rs = Con.OpenRecordSet(Sql)

Dim MyEcel As New EXCEL.Application
Dim MyClasseur As EXCEL.Workbook
Dim MySheet As EXCEL.Worksheet
Dim Myrange As EXCEL.Range

If Fso.FileExists(App.Path & "\ExportPrix\" & XlsPrix & ".xls") = True Then
Fso.DeleteFile App.Path & "\ExportPrix\" & XlsPrix & ".xls"
End If

'MyEcel.Visible = True

Set MyClasseur = MyEcel.Workbooks.Add(App.Path & "\ModèlePrix\ModèlePrix.xlt")

While Rs.EOF = False
i = i + 1
    Rs.MoveNext
Wend
If i = 0 Then i = 1
FormBarGrah.ProgressBar1.Max = i
Rs.Requery
i = 1
While Rs.EOF = False
i = i + 1
    IncremanteBarGrah FormBarGrah
    DoEvents
     MyClasseur.Worksheets("Prix").Cells(i, 1) = Val(Replace("" & Rs!Section, ",", "."))
     MyClasseur.Worksheets("Prix").Cells(i, 2) = "" & Rs!ISO
     MyClasseur.Worksheets("Prix").Cells(i, 3) = Val(Replace("" & Rs![Prix u], ",", "."))
    Rs.MoveNext
Wend
MyClasseur.SaveAs App.Path & "\ExportPrix\" & XlsPrix & ".xls"
FormBarGrah.ProgressBar1Caption.Caption = " Fin du Tratement:"
FormBarGrah.ProgressBar1.Value = 0
DoEvents

MyClasseur.Close False
Set MyClasseur = Nothing
MyEcel.Quit
Set MyEcel = Nothing
End Sub

