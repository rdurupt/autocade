Attribute VB_Name = "ExoprtExcel"
Public MyExcel As EXCEL.Application
Public MyWorkbook As EXCEL.Workbook


Public Sub ExporteXls(Xls As String, IdIndiceProjet As Long)
Dim Fso As New FileSystemObject
Dim sql As String
Dim RsIdProjet As Recordset
Dim Rs As Recordset
Dim PathModelXls As String
 Set TableauPath = funPath

PathModelXls = TableauPath.Item("PathModelXls")
         If Left(PathModelXls, 2) <> "\\" Then PathModelXls = TableauPath.Item("PathServer") & PathModelXls

    Set MyWorkbook = OpenModelXlt(PathModelXls)
    sql = "SELECT Ligne_Tableau_fils.* FROM Ligne_Tableau_fils WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & IdIndiceProjet & " ORDER BY val(Ligne_Tableau_fils.FIL);"
    Set Rs = Con.OpenRecordSet(sql)
    ExporteXlsFils Rs
    
    sql = "SELECT Connecteurs.* FROM Connecteurs "
    sql = sql & "WHERE Connecteurs.Id_IndiceProjet = " & IdIndiceProjet & " "
    sql = sql & "ORDER BY Connecteurs.N°;"
    Set Rs = Con.OpenRecordSet(sql)

    ExporteXlsConnecteur Rs

sql = "SELECT Connecteurs.* FROM Connecteurs "
    sql = sql & "WHERE Connecteurs.Id_IndiceProjet = " & IdIndiceProjet & " "
    sql = sql & "ORDER BY Connecteurs.N°;"
    
    sql = "SELECT Composants.*  "
    sql = sql & "FROM Composants "
    sql = sql & "WHERE Composants.Id_IndiceProjet = " & IdIndiceProjet & " "
    sql = sql & "ORDER BY Composants.NUMCOMP;"
    Set Rs = Con.OpenRecordSet(sql)

    ExporteXlsComposants Rs
    
    sql = "SELECT Connecteurs.* FROM Connecteurs "
    sql = sql & "WHERE Connecteurs.Id_IndiceProjet = " & IdIndiceProjet & " "
    sql = sql & "ORDER BY Connecteurs.N°;"
    sql = "SELECT Nota.* FROM Nota "
    sql = sql & "WHERE Nota.Id_IndiceProjet= " & IdIndiceProjet & " "
    sql = sql & "ORDER BY Nota.NUMNOTA ;"

    Set Rs = Con.OpenRecordSet(sql)

    ExporteXlsNotas Rs

If Fso.FileExists(Xls & ".xls") Then Fso.DeleteFile Xls & ".xls"
Set Fso = Nothing
On Error Resume Next

MyWorkbook.SaveAs Xls
If Err Then MsgBox Err.Description
On Error GoTo 0
MyWorkbook.Close False
Set MyWorkbook = Nothing
MyExcel.Quit

Set MyExcel = Nothing
 FormBarGrah.ProgressBar1.Value = 0
 FormBarGrah.ProgressBar1Caption = "Fin du traitement:"

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
 FormBarGrah.ProgressBar1.Value = 0
 FormBarGrah.ProgressBar1.Max = NbLigne
 FormBarGrah.ProgressBar1Caption.Caption = "Exporter liste des Fils :"
While Rs.EOF = False
     FormBarGrah.ProgressBar1.Value = FormBarGrah.ProgressBar1.Value + 1
    DoEvents
    MyRange(Row, 1) = "'" & Rs!Liai
    MyRange(Row, 2) = "'" & Rs!Designation
    MyRange(Row, 3) = "'" & Rs!Fil
    MyRange(Row, 4) = "'" & Rs!SECT
    MyRange(Row, 5) = "'" & Rs!TEINT
    MyRange(Row, 6) = "'" & Rs!TEINT2
    MyRange(Row, 7) = "'" & Rs!ISO
    MyRange(Row, 8) = "'" & Rs!Long
    MyRange(Row, 9) = "'" & Rs![LONG CP]
    MyRange(Row, 10) = "'" & Rs!Coupe
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
    MyRange(Row, 22) = "'" & Rs![Option]
    Rs.MoveNext
    Row = Row + 1
Wend
Set MyRange = Nothing
Set MySeet = Nothing
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
 FormBarGrah.ProgressBar1.Value = 0
 FormBarGrah.ProgressBar1.Max = NbLigne
 FormBarGrah.ProgressBar1Caption.Caption = "Exporter liste des Connecteurs :"
While Rs.EOF = False
 FormBarGrah.ProgressBar1.Value = FormBarGrah.ProgressBar1.Value + 1
DoEvents
  MyRange(Row, 1) = "'" & Rs!Connecteur
  MyRange(Row, 2) = Abs(Rs![O/N])
  MyRange(Row, 3) = "'" & Rs!Designation
  MyRange(Row, 4) = "'" & Rs!Code_APP
  MyRange(Row, 5) = "'" & Rs![N°]
  MyRange(Row, 6) = "'" & Rs!POS
  MyRange(Row, 7) = "'" & Rs![POS-OUT]
  MyRange(Row, 8) = "'" & Rs!PRECO1
  MyRange(Row, 9) = "'" & Rs!PRECO2
    MyRange(Row, 10) = "'" & Rs![100%]
    Rs.MoveNext
    Row = Row + 1
Wend
Set MyRange = Nothing
Set MySeet = Nothing
End Function
Function ExporteXlsNotas(Rs As Recordset)
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
Set MySeet = MyWorkbook.Worksheets("Notas")
Set MyRange = MySeet.Range("A1").CurrentRegion
Row = 2
 FormBarGrah.ProgressBar1.Value = 0
 FormBarGrah.ProgressBar1.Max = NbLigne
 FormBarGrah.ProgressBar1Caption.Caption = "Exporter liste des Connecteurs :"
While Rs.EOF = False
 FormBarGrah.ProgressBar1.Value = FormBarGrah.ProgressBar1.Value + 1
DoEvents
  MyRange(Row, 1) = "'" & Rs!Nota
  MyRange(Row, 2) = "'" & Rs!NUMNOTA
 
    Rs.MoveNext
    Row = Row + 1
Wend
Set MyRange = Nothing
Set MySeet = Nothing
End Function
Function ExporteXlsComposants(Rs As Recordset)
Dim MySeet As EXCEL.Worksheet
Dim MyRange As EXCEL.Range
Dim Row As Long
Dim Rep As String
Dim NbColonne As Long
Dim TxtPoin As String
Dim NbLigne As Long
Dim sql As String
Dim Fso As New FileSystemObject
Dim PathComposantsDefault As String
Dim RsComposants As Recordset
    NbLigne = 0
Set MySeet = MyWorkbook.Worksheets("Composants")
Set MyRange = MySeet.Range("A1").CurrentRegion
NbColonne = MyRange.Columns.Count
 If Rs.EOF = False Then

     sql = "SELECT T_indiceProjet.Client FROM T_indiceProjet "
    sql = sql & "WHERE T_indiceProjet.Id=" & Rs!Id_IndiceProjet & ";"

    NumErr = 1
    
      sql = "SELECT T_indiceProjet.Client FROM T_indiceProjet "
    sql = sql & "WHERE T_indiceProjet.Id=" & Rs!Id_IndiceProjet & ";"

    NumErr = 1

    Set RsComposants = Con.OpenRecordSet(sql)
    LeCient = UCase(Trim("" & RsComposants!Client))
    sql = "SELECT  T_Clients.PathComposants FROM T_Clients "
sql = sql & "WHERE T_Clients.Client='" & MyReplace(LeCient) & "';"

       Set RsComposants = Con.OpenRecordSet(sql)
If RsComposants.EOF = False Then
    
    If Trim("" & RsComposants!PathComposants) = "" Then
         PathComposantsDefault = TableauPath.Item("PathComposantsDefault")
   Else
             PathComposantsDefault = RsComposants!PathComposants
'         If Left(PathComposantsDefault, 2) <> "\\" Then PathComposantsDefault = TableauPath.Item("PathServer") & PathComposantsDefault
    
    End If
Else
                 PathComposantsDefault = RsComposants!PathComposants

End If
Else
 PathComposantsDefault = TableauPath.Item("PathComposantsDefault")
End If
If Left(PathComposantsDefault, 2) <> "\\" Then PathComposantsDefault = TableauPath.Item("PathServer") & PathComposantsDefault

    Dim fs, f, f1, s, sf
  MyRange(1, NbColonne).AutoFilter
'  MyExcel.Visible = True
    Set f = Fso.GetFolder(PathComposantsDefault) '\\10.30.0.5\\donnees d entreprise\Utilitaires\cablage\Librairies\COMPOSANTS\")
    Set sf = f.SubFolders
    For Each f1 In sf
       NbColonne = NbColonne + 1
    MyRange(1, NbColonne) = f1.Name
    MyRange(1, NbColonne).Interior.ColorIndex = 15
    Next
  
MyRange(1, NbColonne).AutoFilter

If Rs.EOF = True Then Exit Function



TxtPoin = ""
'MyRange(1, NbColonne).AutoFilter
'While Trim(Rep) <> ""
'
'If InStr(1, Trim(Rep), ".") = 0 Then
'    NbColonne = NbColonne + 1
'    MyRange(1, NbColonne) = Rep
'    MyRange(1, NbColonne).Interior.ColorIndex = 15
'
' End If
'    Rep = Dir
'Wend

NbLigne = 1
    If Rs.EOF = True Then Exit Function
While Rs.EOF = False
NbLigne = NbLigne + 1

Rs.MoveNext
Wend
Rs.MoveFirst


Row = 2
 FormBarGrah.ProgressBar1.Value = 0
 FormBarGrah.ProgressBar1.Max = NbLigne
 FormBarGrah.ProgressBar1Caption.Caption = "Exporter liste des Composants :"
 Set MyRange = MySeet.Range("A1").CurrentRegion

While Rs.EOF = False
 FormBarGrah.ProgressBar1.Value = FormBarGrah.ProgressBar1.Value + 1
DoEvents
  MyRange(Row, 1) = "'" & Rs!DESIGNCOMP
  MyRange(Row, 2) = "'" & Rs!NUMCOMP
  MyRange(Row, 3) = "'" & Rs!REFCOMP
  For i = 4 To MyRange.Columns.Count
            If MyRange(1, i) = "" & Rs!Path Then MyRange(Row, i) = 1 Else MyRange(Row, i) = 0
        Next i
  
 
    Row = Row + 1
    Rs.MoveNext
Wend
Set MyRange = Nothing
Set MySeet = Nothing
End Function
Function OpenModelXlt(Fichier As String) As EXCEL.Workbook
Set MyExcel = New EXCEL.Application
'MyExcel.Visible = True
Set OpenModelXlt = MyExcel.Workbooks.Open(Fichier)
End Function

