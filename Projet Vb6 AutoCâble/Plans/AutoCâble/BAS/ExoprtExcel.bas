Attribute VB_Name = "ExoprtExcel"
Public MyExcel As EXCEL.Application
Public MyWorkbook As EXCEL.Workbook


Public Sub ExporteXls(Xls As String, Projet As String, LI As String)
Dim Fso As New FileSystemObject
Dim Sql As String
Dim RsIdProjet As Recordset
Dim Rs As Recordset
Con.OpenConnetion db
 Set TableauPath = funPath
Sql = "SELECT T_indiceProjet.Id "
Sql = Sql & "FROM T_Projet INNER JOIN T_indiceProjet ON T_Projet.id = T_indiceProjet.IdProjet "
Sql = Sql & "WHERE T_Projet.Projet='" & varProjet & "' "
Sql = Sql & "AND T_indiceProjet.LI= '" & MyReplace(varIndice) & "';"
Con.OpenConnetion db
Set RsIdProjet = Con.OpenRecordSet(Sql)
If RsIdProjet.EOF = False Then
    Set MyWorkbook = OpenModelXlt(TableauPath.Item("PathModelXls") & "Ligne_Tableau_fils.xlt")
    Sql = "SELECT Ligne_Tableau_fils.* FROM Ligne_Tableau_fils WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & RsIdProjet!Id & " ORDER BY val(Ligne_Tableau_fils.FIL);"
    Set Rs = Con.OpenRecordSet(Sql)
    ExporteXlsFils Rs
    Sql = "SELECT Connecteurs.* FROM Connecteurs "
    Sql = Sql & "WHERE Connecteurs.Id_IndiceProjet = " & RsIdProjet!Id & " "
    Sql = Sql & "ORDER BY Connecteurs.N°;"
    Set Rs = Con.OpenRecordSet(Sql)

    ExporteXlsConnecteur Rs
End If
Set RsIdProjet = Con.CloseRecordSet(RsIdProjet)
Con.CloseConnection
If Fso.FileExists(Xls) Then Fso.DeleteFile Xls
Set Fso = Nothing
On Error Resume Next
MyWorkbook.SaveAs Xls
If Err Then MsgBox Err.Description
On Error GoTo 0
MyWorkbook.Close False
MyExcel.Quit
Set MyWorkbook = Nothing
Set MyExcel = Nothing
Menu.ProgressBar1.Value = 0
Menu.ProgressBar1Caption = "Fin du traitement:"
Con.CloseConnection
End Sub

Function ExporteXlsFils(Rs As Recordset)
Dim MySeet As EXCEL.Worksheet
Dim MyRange As EXCEL.Range
Dim Row As Long
Dim NbLigne As Long
DoEvents
    NbLigne = 0
If Rs.EOF = True Then Exit Function
While Rs.EOF = False
NbLigne = NbLigne + 1
Rs.MoveNext
Wend
Rs.MoveFirst
Set MySeet = MyWorkbook.Worksheets("Ligne_Tableau_fils")
Set MyRange = MySeet.Range("A1").CurrentRegion
Row = 2
Menu.ProgressBar1.Value = 0
Menu.ProgressBar1.Max = NbLigne
Menu.ProgressBar1Caption.Caption = "Exporter liste des Fils :"
While Rs.EOF = False
    Menu.ProgressBar1.Value = Menu.ProgressBar1.Value + 1
    DoEvents
    MyRange(Row, 1) = "'" & Rs!LIAI
    MyRange(Row, 2) = "'" & Rs!Designation
    MyRange(Row, 3) = "'" & Rs!Fil
    MyRange(Row, 4) = "'" & Rs!SECT
    MyRange(Row, 5) = "'" & Rs!TEINT
    MyRange(Row, 6) = "'" & Rs!TEINT2
    MyRange(Row, 7) = "'" & Rs!ISO
    MyRange(Row, 8) = "'" & Rs!LONG
    MyRange(Row, 9) = "'" & Rs![LONG CP]
    MyRange(Row, 10) = "'" & Rs!COUPE
    MyRange(Row, 11) = "'" & Rs!POS
    MyRange(Row, 12) = "'" & Rs![POS-OUT]
    MyRange(Row, 13) = "'" & Rs!FA
    MyRange(Row, 14) = "'" & Rs![APP]
    MyRange(Row, 15) = "'" & Rs!VOI
    
    MyRange(Row, 16) = "'" & Rs![POS2]
    MyRange(Row, 17) = "'" & Rs![POS-OUT2]
    MyRange(Row, 18) = "'" & Rs![FA2]
    
    MyRange(Row, 19) = "'" & Rs![APP2]
    MyRange(Row, 20) = "'" & Rs![VOI2]
    MyRange(Row, 21) = "'" & Rs![PRECO]
    MyRange(Row, 22) = "'" & Rs![OPTION]
    Rs.MoveNext
    Row = Row + 1
Wend
End Function
Function ExporteXlsConnecteur(Rs As Recordset)
Dim MySeet As EXCEL.Worksheet
Dim MyRange As EXCEL.Range
Dim Row As Long
Dim NbLigne As Long
    NbLigne = 0
    If Rs.EOF = True Then Exit Function
While Rs.EOF = False
NbLigne = NbLigne + 1
Rs.MoveNext
Wend
Rs.MoveFirst
Set MySeet = MyWorkbook.Worksheets("Connecteurs")
Set MyRange = MySeet.Range("A1").CurrentRegion
Row = 2
Menu.ProgressBar1.Value = 0
Menu.ProgressBar1.Max = NbLigne
Menu.ProgressBar1Caption.Caption = "Exporter liste des Connecteurs :"
While Rs.EOF = False
Menu.ProgressBar1.Value = Menu.ProgressBar1.Value + 1
DoEvents
  MyRange(Row, 1) = "'" & Rs!Connecteur
  MyRange(Row, 2) = "'" & Rs![O/N]
  MyRange(Row, 3) = "'" & Rs!Designation
  MyRange(Row, 4) = "'" & Rs!CODE_APP
  MyRange(Row, 5) = "'" & Rs![N°]
  MyRange(Row, 6) = "'" & Rs!POS
  MyRange(Row, 7) = "'" & Rs![POS-OUT]
  MyRange(Row, 8) = "'" & Rs!PRECO1
  MyRange(Row, 9) = "'" & Rs!PRECO2
    MyRange(Row, 10) = "'" & Rs![100%]
    Rs.MoveNext
    Row = Row + 1
Wend
End Function
Function OpenModelXlt(Fichier As String) As EXCEL.Workbook
Set MyExcel = New EXCEL.Application
'MyExcel.Visible = True

Set OpenModelXlt = MyExcel.Workbooks.Open(Fichier)
End Function

