Attribute VB_Name = "ImporterXls"
Public Sub ImporteXls(Xls As String, IdIndiceProjet As Long)
Dim sql As String
Dim Rs As Recordset
Dim IdProjet As Long
Dim IdIndice As Long

DoEvents

Set TableauPath = funPath
IdIndice = IdIndiceProjet

Dim MyEcel As New EXCEL.Application
Dim MyClasseur As EXCEL.Workbook
Dim MySheet As EXCEL.Worksheet
Dim MyRange As EXCEL.Range
Set MyClasseur = MyEcel.Workbooks.Open(Xls)
Set MySheet = MyClasseur.Worksheets("Ligne_Tableau_fils")

sql = "DELETE xls_Ligne_Tableau_fils.*  FROM xls_Ligne_Tableau_fils "
sql = sql & "where xls_Ligne_Tableau_fils.Job=" & NmJob & ";"
Con.Exequte sql

sql = "DELETE Xls_Connecteurs.* FROM Xls_Connecteurs "
sql = sql & "WHERE Xls_Connecteurs.Job=" & NmJob & ";"
Con.Exequte sql

sql = "DELETE Xls_Nota.* FROM Xls_Nota "
sql = sql & "WHERE Xls_Nota.Job=" & NmJob & ";"
Con.Exequte sql

sql = "DELETE Xls_Composants.* FROM Xls_Composants "
sql = sql & "WHERE Xls_Composants.Job=" & Xls_Composants & ";"
Con.Exequte sql
'MyEcel.Visible = True
Set MyRange = MySheet.Range("A1").CurrentRegion
 FormBarGrah.ProgressBar1.Value = 0
 FormBarGrah.ProgressBar1.Max = MyRange.Rows.Count
 FormBarGrah.ProgressBar1Caption.Caption = "Importe la liste de fils"

For Row = 2 To MyRange.Rows.Count
 FormBarGrah.ProgressBar1.Value = FormBarGrah.ProgressBar1.Value + 1
DoEvents

   sql = sqlRange(MyRange, Row, "Xls_Ligne_Tableau_fils")
   Con.Exequte sql
Next Row

Set MySheet = MyClasseur.Worksheets("Connecteurs")
Set MyRange = MySheet.Range("A1").CurrentRegion
 FormBarGrah.ProgressBar1.Value = 0
 FormBarGrah.ProgressBar1.Max = MyRange.Rows.Count
 FormBarGrah.ProgressBar1Caption.Caption = "Importe la liste des Connecteurs"



For Row = 2 To MyRange.Rows.Count
 FormBarGrah.ProgressBar1.Value = FormBarGrah.ProgressBar1.Value + 1
   sql = sqlRange(MyRange, Row, "Xls_Connecteurs")
   Con.Exequte sql
Next Row



Set MySheet = MyClasseur.Worksheets("Composants")
Set MyRange = MySheet.Range("A1").CurrentRegion
 FormBarGrah.ProgressBar1.Value = 0
 FormBarGrah.ProgressBar1.Max = MyRange.Rows.Count
 FormBarGrah.ProgressBar1Caption.Caption = "Importe la liste des Composants"



For Row = 2 To MyRange.Rows.Count
 FormBarGrah.ProgressBar1.Value = FormBarGrah.ProgressBar1.Value + 1
   sql = sqlRange(MyRange, Row, "Xls_Composants")
   Con.Exequte sql
Next Row


Set MySheet = MyClasseur.Worksheets("Notas")
Set MyRange = MySheet.Range("A1").CurrentRegion
 FormBarGrah.ProgressBar1.Value = 0
 FormBarGrah.ProgressBar1.Max = MyRange.Rows.Count
 FormBarGrah.ProgressBar1Caption.Caption = "Importe la liste des Notas"



For Row = 2 To MyRange.Rows.Count
 FormBarGrah.ProgressBar1.Value = FormBarGrah.ProgressBar1.Value + 1
   sql = sqlRange(MyRange, Row, "Xls_Nota")
   Con.Exequte sql
Next Row



Set MyRange = Nothing
Set MySheet = Nothing
MyClasseur.Close False
Set MyClasseur = Nothing
MyEcel.Quit
Set MyEcel = Nothing
MajBase IdIndice
sql = "DELETE xls_Ligne_Tableau_fils.*  FROM xls_Ligne_Tableau_fils "
sql = sql & "where xls_Ligne_Tableau_fils.Job=" & NmJob & ";"
Con.Exequte sql

sql = "DELETE Xls_Connecteurs.* FROM Xls_Connecteurs "
sql = sql & "WHERE Xls_Connecteurs.Job=" & NmJob & ";"
Con.Exequte sql

sql = "DELETE Xls_Nota.* FROM Xls_Nota "
sql = sql & "WHERE Xls_Nota.Job=" & NmJob & ";"
Con.Exequte sql

sql = "DELETE Xls_Composants.* FROM Xls_Composants "
sql = sql & "WHERE Xls_Composants.Job=" & Xls_Composants & ";"
Con.Exequte sql

 FormBarGrah.ProgressBar1.Value = 0
 FormBarGrah.ProgressBar1Caption.Caption = "Fin du traitement"
End Sub


Function sqlRange(MyRange As EXCEL.Range, Row, FROM)
Dim Sql1 As String
Dim Sql2 As String
Dim Sql1Val As String
Dim Sql2Val As String
Dim Sql3Val As String
Sql1 = "INSERT INTO " & FROM & " (Job,"
Sql1Val = ""
Sql2Val = NmJob & ","
Sql3Val = ""
For i = 1 To MyRange.Columns.Count
     If FROM = "Xls_Composants" And i > 3 Then Exit For
DoEvents
    Sql1Val = Sql1Val & "[" & MyRange(1, i) & "],"
   
    If Trim("" & MyRange(Row, i)) = "" Then
    If MyRange(1, i) = "O/N" Then
         Sql2Val = Sql2Val & "0,"
    Else
        Sql2Val = Sql2Val & "NULL,"
    End If
    Else
        If MyRange(1, i) = "O/N" Then
        If UCase(Trim(MyRange(Row, i))) = "N" Then MyRange(Row, i) = 0
         If UCase(Trim(MyRange(Row, i))) = "O" Then MyRange(Row, i) = 1
            Sql2Val = Sql2Val & "" & CInt(Trim(MyRange(Row, i))) & ","
        Else
            Sql2Val = Sql2Val & "'" & MyReplace(Trim(MyRange(Row, i))) & "',"
        End If
    End If
   
Next i
If FROM = "Xls_Composants" Then
Sql1Val = Sql1Val & "[Path],"
Sql3Val = "NULL,"
    For i = i To MyRange.Columns.Count
        If Val((Trim("" & MyRange(Row, i).Value))) = 1 Then
            Sql3Val = "'" & MyReplace(Trim(MyRange(1, i))) & "',"
            Exit For
        End If
    Next i
End If

Sql1Val = Left(Sql1Val, Len(Sql1Val) - 1)
Sql2Val = Sql2Val & Sql3Val
Sql2Val = Left(Sql2Val, Len(Sql2Val) - 1)
sqlRange = Sql1 & Sql1Val & ") Values(" & Sql2Val & ");"
End Function

Public Sub CeerFichierXls(Xls As String)
Dim sql As String
Dim Rs As Recordset
Dim IdProjet As Long
Dim IdIndice As Long
DoEvents

Dim MyEcel As New EXCEL.Application
Dim MyClasseur As EXCEL.Workbook
Dim MySheet As EXCEL.Worksheet
Dim MyRange As EXCEL.Range
Set MyClasseur = MyEcel.Workbooks.Open(TableauPath.Item("PathModelXls") & "Ligne_Tableau_fils.xlt")
MyEcel.DisplayAlerts = False
MyClasseur.SaveAs Xls
MyEcel.DisplayAlerts = True
MyEcel.Visible = True
End Sub
