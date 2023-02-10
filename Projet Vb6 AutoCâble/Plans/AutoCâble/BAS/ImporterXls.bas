Attribute VB_Name = "ImporterXls"
Public Sub ImporteXls(Xls As String, Projet As String, Indce As String, Description As String, LI As String, Cle_Ch)
Dim Sql As String
Dim Rs As Recordset
Dim IdProjet As Long
Dim IdIndice As Long

DoEvents
Con.OpenConnetion db
Set TableauPath = funPath
Sql = "SELECT T_Projet.id FROM T_Projet WHERE T_Projet.Projet='" & MyReplace(Projet) & "';"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = False Then
Sql = "UPDATE T_Projet SET T_Projet.CleAc = " & Cle_Ch & " WHERE T_Projet.id=" & Rs!Id & ";"
Con.Exequte Sql

  IdProjet = Rs!Id
Else

    Sql = "INSERT INTO T_Projet ( Projet,CleAc)"
    Sql = Sql & "Values('" & MyReplace(Projet) & "'," & Cle_Ch & ");"
    Con.Exequte Sql

 
Rs.Requery
 IdProjet = Rs!Id
End If

Sql = "SELECT T_indiceProjet.id FROM T_indiceProjet WHERE T_indiceProjet.Li='" & MyReplace(LI) & "' and T_indiceProjet.IdProjet=" & IdProjet & ";"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = False Then
IdIndice = Rs!Id
Sql = "UPDATE T_indiceProjet SET T_indiceProjet.Li = '" & MyReplace(LI) & "' WHERE T_indiceProjet.Id=" & IdIndice & ";"
Con.Exequte Sql
  
Else
 If EDITER.Caption = "Modifier un plan :" Then
    Sql = "SELECT T_indiceProjet.Id, T_indiceProjet.IdProjet, T_indiceProjet.Indice, T_indiceProjet.Description, T_indiceProjet.Li, T_indiceProjet.IdStatus, T_indiceProjet.IdApprobateur, T_indiceProjet.AutoCadSaveAs, T_indiceProjet.AutoCadSave "
    Sql = Sql & "FROM T_indiceProjet "
    Sql = Sql & "WHERE T_indiceProjet.LI='" & MyReplace(EDITER.lstIndice.List(EDITER.lstIndice.ListIndex, 0)) & "';"
    Set Rs = Con.OpenRecordSet(Sql)
    If Rs.EOF = False Then
    Sql = "INSERT INTO T_indiceProjet ( IdProjet, Indice, Description, Li, IdStatus, IdApprobateur, AutoCadSave ) "
    Sql = Sql & "VALUES(" & Rs!IdProjet & ", '" & Rs!Indice & "','" & Rs!Description & "','" & LI & "', 2," & Rs!IdApprobateur & ",'" & Rs!AutoCadSaveAs & "') "
'    Sql = Sql & "WHERE T_indiceProjet.Id=37;"
Con.Exequte Sql
    Sql = "SELECT T_indiceProjet.Id, T_indiceProjet.IdProjet, T_indiceProjet.Indice, T_indiceProjet.Description, T_indiceProjet.Li, T_indiceProjet.IdStatus, T_indiceProjet.IdApprobateur, T_indiceProjet.AutoCadSaveAs, T_indiceProjet.AutoCadSave "
    Sql = Sql & "FROM T_indiceProjet "
    Sql = Sql & "WHERE T_indiceProjet.LI='" & MyReplace(LI) & "';"
    Set Rs = Con.OpenRecordSet(Sql)

    End If
 Else
Sql = "INSERT INTO T_indiceProjet ( IdProjet, Indice, Description ,Li)"
Sql = Sql & "values( " & IdProjet & " , '" & MyReplace(Indce) & "', '" & MyReplace(Description) & "','" & MyReplace(LI) & "' );"
Con.Exequte Sql
  Sql = "SELECT T_indiceProjet.Id, T_indiceProjet.IdProjet, T_indiceProjet.Indice, T_indiceProjet.Description, T_indiceProjet.Li, T_indiceProjet.IdStatus, T_indiceProjet.IdApprobateur, T_indiceProjet.AutoCadSaveAs, T_indiceProjet.AutoCadSave "
    Sql = Sql & "FROM T_indiceProjet "
    Sql = Sql & "WHERE T_indiceProjet.LI='" & MyReplace(LI) & "';"
    Set Rs = Con.OpenRecordSet(Sql)

End If
IdIndice = Rs!Id
End If

Dim MyEcel As New EXCEL.Application
Dim MyClasseur As EXCEL.Workbook
Dim MySheet As EXCEL.Worksheet
Dim MyRange As EXCEL.Range
'MyEcel.Visible = True
Set MyClasseur = MyEcel.Workbooks.Open(Xls)
Set MySheet = MyClasseur.Worksheets("Ligne_Tableau_fils")

Sql = "DELETE xls_Ligne_Tableau_fils.*  FROM xls_Ligne_Tableau_fils;"
Con.Exequte Sql
Sql = "DELETE Xls_Connecteurs.* FROM Xls_Connecteurs;"
Con.Exequte Sql
Set MyRange = MySheet.Range("A1").CurrentRegion
Menu.ProgressBar1.Value = 0
Menu.ProgressBar1.Max = MyRange.Rows.Count
Menu.ProgressBar1Caption.Caption = "Importe la liste de fils"

For Row = 2 To MyRange.Rows.Count
Menu.ProgressBar1.Value = Menu.ProgressBar1.Value + 1
DoEvents

   Sql = sqlRange(MyRange, Row, "Xls_Ligne_Tableau_fils")
   Con.Exequte Sql
Next Row

Set MySheet = MyClasseur.Worksheets("Connecteurs")
Set MyRange = MySheet.Range("A1").CurrentRegion
Menu.ProgressBar1.Value = 0
Menu.ProgressBar1.Max = MyRange.Rows.Count
Menu.ProgressBar1Caption.Caption = "Importe la liste des Connecteurs"



For Row = 2 To MyRange.Rows.Count
Menu.ProgressBar1.Value = Menu.ProgressBar1.Value + 1
   Sql = sqlRange(MyRange, Row, "Xls_Connecteurs")
   Con.Exequte Sql
Next Row
Set MyRange = Nothing
Set MySheet = Nothing
MyClasseur.Close False
Set MyClasseur = Nothing
MyEcel.Quit
Set MyEcel = Nothing
MajBase IdIndice
Con.CloseConnection
Menu.ProgressBar1.Value = 0
Menu.ProgressBar1Caption.Caption = "Fin du traitement"
End Sub

Sub test()
    ImporteXls "C:\RD\FichierXls\Ligne_Tableau_fils.xls"
    
End Sub
Sub MajBase(IdIndice As Long)
Dim Sql As String
Sql = "DELETE Ligne_Tableau_fils.* "
Sql = Sql & "FROM Ligne_Tableau_fils "
Sql = Sql & "WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & IdIndice & ";"
Con.Exequte Sql

Sql = "DELETE Connecteurs.* "
Sql = Sql & "FROM Connecteurs "
Sql = Sql & "WHERE Connecteurs.Id_IndiceProjet=" & IdIndice & ";"
Con.Exequte Sql

Sql = "INSERT INTO Ligne_Tableau_fils ( Id_IndiceProjet, LIAI, "
Sql = Sql & "DESIGNATION, FIL, SECT, TEINT, TEINT2, ISO, POS, FA,  "
Sql = Sql & "VOI, POS2, FA2, VOI2, [LONG] ,APP,APP2,PRECO,[OPTION]) "
Sql = Sql & "SELECT " & IdIndice & " AS Id_IndiceProjet, xls_Ligne_Tableau_fils.LIAI,  "
Sql = Sql & "xls_Ligne_Tableau_fils.DESIGNATION, xls_Ligne_Tableau_fils.FIL,  "
Sql = Sql & "xls_Ligne_Tableau_fils.SECT, xls_Ligne_Tableau_fils.TEINT,  "
Sql = Sql & "xls_Ligne_Tableau_fils.TEINT2, xls_Ligne_Tableau_fils.ISO,  "
Sql = Sql & "xls_Ligne_Tableau_fils.POS, xls_Ligne_Tableau_fils.FA,  "
Sql = Sql & "xls_Ligne_Tableau_fils.VOI, xls_Ligne_Tableau_fils.POS2,  "
Sql = Sql & "xls_Ligne_Tableau_fils.FA2, xls_Ligne_Tableau_fils.VOI2,  "
Sql = Sql & "xls_Ligne_Tableau_fils.LONG,xls_Ligne_Tableau_fils.APP as APP,xls_Ligne_Tableau_fils.APP2 as APP2,xls_Ligne_Tableau_fils.PRECO as PRECO,xls_Ligne_Tableau_fils.[OPTION] as [OPTION]"
Sql = Sql & "FROM xls_Ligne_Tableau_fils;"


Sql = "INSERT INTO Ligne_Tableau_fils ( Id_IndiceProjet, LIAI, DESIGNATION,  "
Sql = Sql & "FIL, SECT, TEINT, TEINT2, ISO, [LONG], [LONG CP], COUPE, POS,  "
Sql = Sql & "[POS-OUT], FA, APP, VOI, POS2, [POS-OUT2], FA2, APP2,  "
Sql = Sql & "VOI2, PRECO, [OPTION] ) "
Sql = Sql & "SELECT " & IdIndice & " AS Id_IndiceProjet, xls_Ligne_Tableau_fils.LIAI,  "
Sql = Sql & "xls_Ligne_Tableau_fils.DESIGNATION, xls_Ligne_Tableau_fils.FIL,  "
Sql = Sql & "xls_Ligne_Tableau_fils.SECT, xls_Ligne_Tableau_fils.TEINT,  "
Sql = Sql & "xls_Ligne_Tableau_fils.TEINT2, xls_Ligne_Tableau_fils.ISO,  "
Sql = Sql & "xls_Ligne_Tableau_fils.LONG, xls_Ligne_Tableau_fils.[LONG CP],  "
Sql = Sql & "xls_Ligne_Tableau_fils.COUPE, xls_Ligne_Tableau_fils.POS,  "
Sql = Sql & "xls_Ligne_Tableau_fils.[POS-OUT], xls_Ligne_Tableau_fils.FA,  "
Sql = Sql & "xls_Ligne_Tableau_fils.APP, xls_Ligne_Tableau_fils.VOI,  "
Sql = Sql & "xls_Ligne_Tableau_fils.POS2, xls_Ligne_Tableau_fils.[POS-OUT2],  "
Sql = Sql & "xls_Ligne_Tableau_fils.FA2, xls_Ligne_Tableau_fils.APP2,  "
Sql = Sql & "xls_Ligne_Tableau_fils.VOI2, xls_Ligne_Tableau_fils.PRECO,  "
Sql = Sql & "xls_Ligne_Tableau_fils.OPTION "
Sql = Sql & "FROM xls_Ligne_Tableau_fils;"




Con.Exequte Sql

Sql = "INSERT INTO Connecteurs ( Id_IndiceProjet, CONNECTEUR,  "
Sql = Sql & "[EPISSURE O/N], DESIGNATION, CODE_APP, N°, POS,  "
Sql = Sql & "PRECO1, PRECO2 ) "
Sql = Sql & "SELECT " & IdIndice & " AS Id_IndiceProjet,Xls_Connecteurs.CONNECTEUR, "
Sql = Sql & "Xls_Connecteurs.[EPISSURE O/N], Xls_Connecteurs.DESIGNATION,  "
Sql = Sql & "Xls_Connecteurs.CODE_APP, Xls_Connecteurs.N°, "
Sql = Sql & "Xls_Connecteurs.POS, Xls_Connecteurs.PRECO1,  "
Sql = Sql & "Xls_Connecteurs.PRECO2 "
Sql = Sql & "FROM Xls_Connecteurs;"

Sql = "INSERT INTO Connecteurs ( Id_IndiceProjet, CONNECTEUR, [O/N],  "
Sql = Sql & "DESIGNATION, CODE_APP, N°, POS, [POS-OUT], PRECO1, PRECO2,  "
Sql = Sql & "[100%] )  "
Sql = Sql & "SELECT " & IdIndice & "  AS Id_IndiceProjet, Xls_Connecteurs.CONNECTEUR,  "
Sql = Sql & "Xls_Connecteurs.[O/N], Xls_Connecteurs.DESIGNATION,  "
Sql = Sql & "Xls_Connecteurs.CODE_APP, Xls_Connecteurs.N°,  "
Sql = Sql & "Xls_Connecteurs.POS, Xls_Connecteurs.[POS-OUT],  "
Sql = Sql & "Xls_Connecteurs.PRECO1, Xls_Connecteurs.PRECO2,  "
Sql = Sql & "Xls_Connecteurs.[100%] "
Sql = Sql & "FROM Xls_Connecteurs;"




Con.Exequte Sql

End Sub
Function sqlRange(MyRange As EXCEL.Range, Row, FROM)
Dim Sql1 As String
Dim Sql2 As String
Dim Sql1Val As String
Dim Sql2Val As String
Sql1 = "INSERT INTO " & FROM & " ("
Sql1Val = ""
Sql2Val = ""
For i = 1 To MyRange.Columns.Count
DoEvents
    Sql1Val = Sql1Val & "[" & MyRange(1, i) & "],"
   
    If Trim("" & MyRange(Row, i)) = "" Then
        Sql2Val = Sql2Val & "NULL,"
    Else
     Sql2Val = Sql2Val & "'" & Trim(MyRange(Row, i)) & "',"
    End If
Next i

' LIAI, DESIGNATION, FIL, SECT, TEINT, TEINT2, ISO, POS, FA, VOI, POS2, FA2, VOI2, [LONG] )
'SELECT xls_Ligne_Tableau_fils.LIAI, xls_Ligne_Tableau_fils.DESIGNATION, xls_Ligne_Tableau_fils.FIL, xls_Ligne_Tableau_fils.SECT, xls_Ligne_Tableau_fils.TEINT, xls_Ligne_Tableau_fils.TEINT2, xls_Ligne_Tableau_fils.ISO, xls_Ligne_Tableau_fils.POS, xls_Ligne_Tableau_fils.FA, xls_Ligne_Tableau_fils.VOI, xls_Ligne_Tableau_fils.POS2, xls_Ligne_Tableau_fils.FA2, xls_Ligne_Tableau_fils.VOI2, xls_Ligne_Tableau_fils.LONG
'FROM xls_Ligne_Tableau_fils;
Sql1Val = Left(Sql1Val, Len(Sql1Val) - 1)
Sql2Val = Left(Sql2Val, Len(Sql2Val) - 1)
sqlRange = Sql1 & Sql1Val & ") Values(" & Sql2Val & ");"
End Function

Public Sub CeerFichierXls(Xls As String)
Dim Sql As String
Dim Rs As Recordset
Dim IdProjet As Long
Dim IdIndice As Long
DoEvents
'Con.OpenConnetion db
'Sql = "SELECT T_Projet.id FROM T_Projet WHERE T_Projet.Projet='" & Projet & "';"
'Set Rs = Con.OpenRecordSet(Sql)
'If Rs.EOF = False Then
'  IdProjet = Rs!Id
'Else
'Sql = "INSERT INTO T_Projet ( Projet )"
'Sql = Sql & "Values('" & Projet & "');"
'Con.Exequte Sql
'Rs.Requery
' IdProjet = Rs!Id
'End If
'
'Sql = "SELECT T_indiceProjet.id FROM T_indiceProjet WHERE T_indiceProjet.Indice='" & Indce & "' and T_indiceProjet.IdProjet=" & IdProjet & ";"
'Set Rs = Con.OpenRecordSet(Sql)
'If Rs.EOF = False Then
'  IdIndice = Rs!Id
'Else
'Sql = "INSERT INTO T_indiceProjet ( IdProjet, Indice, Description )"
'Sql = Sql & "values( " & IdProjet & " , '" & Indce & "', '" & Description & "' );"
'Con.Exequte Sql
'Rs.Requery
'IdIndice = Rs!Id
'End If

Dim MyEcel As New EXCEL.Application
Dim MyClasseur As EXCEL.Workbook
Dim MySheet As EXCEL.Worksheet
Dim MyRange As EXCEL.Range
'MyEcel.Visible = True
Set MyClasseur = MyEcel.Workbooks.Open(TableauPath.Item("PathModelXls") & "Ligne_Tableau_fils.xlt")
MyEcel.DisplayAlerts = False
'MyClasseur.de
MyClasseur.SaveAs Xls
MyEcel.DisplayAlerts = True
MyEcel.Visible = True
End Sub
