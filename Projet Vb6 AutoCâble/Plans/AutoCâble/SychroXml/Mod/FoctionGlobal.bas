Attribute VB_Name = "FoctionGlobal"
Option Explicit
Sub IncrmentServer(Optional Mytype As String)
On Error Resume Next

Dim Sql As String

    If Trim("" & Mytype) = "" Then
    
        Sql = "UPDATE T_Job SET Status ='" & MyReplace(FormBarGrah.ProgressBar1Caption) & "' "
    Else
        Sql = "UPDATE T_Job SET Status ='" & MyReplace(Mytype) & " : " & MyReplace(FormBarGrah.ProgressBar1Caption) & "'"
    End If
    Sql = Sql & ",MaxBarGraph = " & FormBarGrah.ProgressBar1.Max & " "
    Sql = Sql & ",ValBarGraph = " & FormBarGrah.ProgressBar1.Value & ", T_Job.BarGraphMaj = Now() "
    Sql = Sql & "WHERE T_Job.Job= " & Command & ";"
    Con.Execute Sql
    Con.Execute Sql
    Con.Execute Sql
Err.Clear


End Sub

 Function AtrbNumError() As Long
    Dim Sql As String
    Dim NErr As Long
    Dim RsNumError As Recordset
    Sql = "SELECT T_NumErreur.LibErreur, T_NumErreur.NumErreur "
    Sql = Sql & "FROM T_NumErreur "
    Sql = Sql & "WHERE T_NumErreur.LibErreur='ErrorApp';"
    Set RsNumError = Con.OpenRecordSet(Sql)
    If RsNumError.EOF = False Then
        Sql = "UPDATE T_NumErreur SET T_NumErreur.NumErreur = [NumErreur]+1;"
        Con.Execute Sql
        RsNumError.Requery
        AtrbNumError = RsNumError!NumErreur
    End If
End Function
Function MsgErreur(NumErr As Long, Lib1 As String, Lib2 As String, ErrDetail As String) As String
    Select Case NumErr
                Case 1
                    MsgErreur = "Le connecteur : " & Lib1 & " Réf : " & Lib2 & " n''existe pas dans la bibliothèque de blocks."
                Case 2
                    MsgErreur = "L''attribut : " & Lib1 & " de : " & Lib2 & " n''existe pas."
                Case 3
                    MsgErreur = "Impossible d''affecter le fil N° : " & Lib1 & " au connecteur : " & Lib2 & " car celui-ci n''existe pas."
                Case 4
                    MsgErreur = "Erreur de numérotation pour le connecteur : " & Lib1 & " vérifiez s''il n''existe pas un trou dans la numérotaion.  "
                Case 5
                    MsgErreur = "L''attribut : " & Lib1 & " du connecteur : " & Lib2 & " n''existe pas."
                Case 6
                    MsgErreur = "Le composant : " & Lib1 & " Réf : " & Lib2 & " n''existe pas dans la bibliothèque de blocks."
                Case 7
                    MsgErreur = "L''attribut : " & Lib1 & " du composant : " & Lib2 & " n''existe pas."
                Case 8
                    MsgErreur = "Le connecteur : " & Lib1 & " n''existe pas dans le catalogue Client."
                Case 9
                    MsgErreur = "Le Block : " & Lib1 & " n''existe pas dans la bibliothèque de blocks."
                Case 10
                    MsgErreur = "Le fichier :  " & Lib1 & vbCrLf & "est actuellement ouvert par un autre utilisateur  et ne peut pas être sauvegardé."
                 Case 11
                    MsgErreur = "Pb Excel :  le fichier EXCEL ne peut pas être enregistré."
                 
    End Select
    MsgErreur = MsgErreur & vbCrLf & "Détail de l''erreur :"
    MsgErreur = MsgErreur & vbCrLf & "********************************************************************************************"
    MsgErreur = MsgErreur & vbCrLf & MyReplace(ErrDetail)
    MsgErreur = MsgErreur & vbCrLf & "********************************************************************************************"
    MsgErreur = MsgErreur & vbCrLf
    MsgErreur = MsgErreur & vbCrLf
    Debug.Print MsgErreur
    NbError = NbError + 1
End Function
Function FunError(NumErr As Long, Lib1 As String, msg As String, Optional Lib2 As String)
Dim Sql As String
If Trim("" & Lib1) = "" Then Exit Function
If JobError = 0 Then JobError = AtrbNumError
msg = MsgErreur(NumErr, Lib1, Lib2, msg)
Sql = "INSERT INTO T_Error ( JobError, ValError ) "
Sql = Sql & "values(" & JobError & ",'" & msg & "' );"
Con.Execute Sql

End Function
Function NoeuName(Row As Long)
Dim I As Long
Dim Txt As String
Dim Ofset As Long
Dim nbTour As Long
Dim NbTord As Long
Dim txtColone As Long
Dim txtNuberColone As Long

Txt = "AA"
txtColone = Len(Txt)
txtNuberColone = Len(Txt)
Ofset = 0
nbTour = 0
NbTord = 0


For I = 0 To Row - 3
Reprise:
Mid(Txt, txtColone, 1) = Chr(Asc(Mid(Txt, txtColone, 1)) + 1)
DoEvents
If Asc(Mid(Txt, txtColone, 1)) = 91 Then
Mid(Txt, txtColone, 1) = "A"
txtColone = txtColone - 1
If txtColone = 0 Then
    Txt = Txt & "A"
    txtColone = Len(Txt)
Else
    GoTo Reprise
End If

End If
   If txtColone <> Len(Txt) Then txtColone = Len(Txt)



Next

NoeuName = Txt
End Function

Public Sub IncremanteBarGrah(obj As Object)
On Error Resume Next
If obj.ProgressBar1.Max = obj.ProgressBar1.Value Then
            obj.ProgressBar1.Max = obj.ProgressBar1.Max + 1
        End If
         obj.ProgressBar1.Value = obj.ProgressBar1.Value + 1
        
         DoEvents
 On Error GoTo 0
End Sub
Sub ExporteFrmCriteresFils(Id_Pieces As Long, MySeet As Object)
Dim Sql As String
Dim Rs As Recordset
Dim MyRange As Object
'Dim MySeet As EXCEL.Worksheet
Dim L As Long
Dim C As Long
Dim Equ As Long
Dim Equ2 As Long
Dim Equipement
Dim Equipement2
Dim Trouve As Boolean
Sql = "SELECT T_indiceProjet.*, T_indiceProjet.Id "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "Where T_indiceProjet.Pere = " & Id_Pieces & "  or T_indiceProjet.Id=" & Id_Pieces & "  "
Sql = Sql & "ORDER BY T_indiceProjet.Id DESC;"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = True Then Exit Sub
Sql = "SELECT  [PI] & '_' & Trim([PL_Indice]) AS Piece "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "Where T_indiceProjet.Id = " & Id_Pieces & " "
Sql = Sql & "Or T_indiceProjet.Pere = " & Id_Pieces & " "
Sql = Sql & "GROUP BY [PI] & '_' & Trim([PL_Indice]) ;"
'Sql = Sql & "ORDER BY T_indiceProjet.Id DESC;"
'Set MySeet = IsertSheet(MyWorkbook, "Critères", True)
'MySeet.Application.Visible = True
Set Rs = Con.OpenRecordSet(Sql)
While Rs.EOF = False
    Set MyRange = MySeet.Range("A1").CurrentRegion
    Set MyRange = MyRange(1, MyRange.Columns.Count + 1)
'    Myrange.Application.Visible = True
    ExcelCreatTitre MyRange, Rs, True, True
    MySeet.Cells(1, MySeet.Range("A1").CurrentRegion.Columns.Count).Interior.Color = ChoixCouleur(0)
    Rs.MoveNext
Wend
Set Rs = Con.CloseRecordSet(Rs)
Sql = "SELECT  T_indiceProjet.Equipement,[PI] & '_' & Trim([PL_Indice]) AS Piece "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "Where T_indiceProjet.Id = " & Id_Pieces & " "
Sql = Sql & "Or T_indiceProjet.Pere = " & Id_Pieces & " "
Sql = Sql & "ORDER BY T_indiceProjet.Id DESC;"
'Myrange.Application.Visible = True
Set MyRange = MySeet.Range("A1").CurrentRegion
Set Rs = Con.OpenRecordSet(Sql)

While Rs.EOF = False
    For C = 1 To MyRange.Columns.Count
        If MyRange(1, C) = "" & Rs!Piece Then Exit For
    Next
    Equipement = Split("" & Rs!Equipement & ";", ";")
    For Equ = 0 To UBound(Equipement) - 1
        Equipement2 = Split("" & Equipement(Equ) & "_", "_")
         
         Set MyRange = MySeet.Range("A1").CurrentRegion
            For L = 2 To MyRange.Rows.Count
            Trouve = False
                If UCase(MyRange(L, 3)) = UCase(Equipement2(0)) Then
                Trouve = True
                If Equipement(Equ) = "" Then Exit For
                    MyRange(L, C) = "X"
                    
             
                    Exit For
                End If
            Next
            If Trouve = False And Equipement(Equ) <> "" Then
            L = MyRange.Rows.Count + 1
                  MyRange(L, C) = "X"
           
                MyRange(L, 1) = 1
                MyRange(L, 2) = UCase(Equipement2(0))
                MyRange(L, 3) = UCase(Equipement2(0))
             
         End If
        
       
    Next
    Rs.MoveNext
Wend

'MyRange.Application.Visible = True
Set MyRange = MySeet.Range("A1").CurrentRegion
MyRange.AutoFilter
MyRange.AutoFilter
On Error Resume Next
    MyRange.CurrentRegion.AutoFitColumns
    Err.Clear
    MyRange.Cells.EntireColumn.AutoFit
    Err.Clear
    On Error GoTo 0

Set MyRange = Nothing
'Trier MySeet, 1, Myrange.Address, "B1", 1, "", 0, "", 0
'sql = "SELECT  T_indiceProjet.Ensemble, T_indiceProjet.Equipement,[Li] & '_' & [LI_Indice] AS Liste, "
'sql = sql & "[PI] & '_' & [PI_Indice] AS Piece, T_indiceProjet.Client, T_indiceProjet.CleAc "
'sql = sql & "FROM T_indiceProjet "
'sql = sql & "WHERE T_indiceProjet.Id=" & Id_Pieces & ";"
'Dim RsEntetePage As Recordset
'
'Set RsEntetePage = Con.OpenRecordSet(sql)
'
'         MyPiedTxt = "Debut : __/__/____" & vbCrLf
'MyPiedTxt = MyPiedTxt & "Fin : __/__/____" & vbCrLf
'MyPiedTxt = MyPiedTxt & "Réalisé par :"
'
'MiseEnPage MySeet, Myrange, "Affaire : " & RsEntetePage!CleAc & Chr(10) & _
' _
'     "Pièce : " & RsEntetePage!Piece & Chr(10) & "Liste : " & RsEntetePage!Liste, vbCrLf & "Câblage : " & Replace("" & RsEntetePage!Ensemble, vbCrLf, " ") & Chr(10) & "Equipement : " & Replace("" & RsEntetePage!Equipement, vbCrLf, " ") & Chr(10) & _
'     "" _
'    , "ENCELADE" & Chr(10) & "Client : " & RsEntetePage!Client & Chr(10) & Format(Date, "dd-mmm-yyyy"), "", "" & MyPiedTxt, "&P/&N", 100, "A2", True, xlPortrait, True, False, False, 2.5, True, True
'
'MaJEncadreXls Myrange, xlThin, xlThin, xlHairline, xlHairline
'
'Set RsEntetePage = Con.CloseRecordSet(RsEntetePage)

End Sub
'Function ExporteXlsPrixFils(Rs As Recordset, Id_IndiceProjet As Long)
'Dim MySeet As Excel.Worksheet
'Dim MyRange As Excel.Range
''Dim MyWorkbook As Workbook
'Dim Row As Long
'Dim NbLigne As Long
'Dim I As Long
'Dim R1 As Long
'Dim R2 As Long
'DoEvents
'    NbLigne = 0
'If Rs.EOF = True Then Exit Function
''While Rs.EOF = False
''NbLigne = NbLigne + 1
''Rs.MoveNext
''Wend
'
'Rs.Requery
''MyWorkbook.Application.Visible = True
'Set MySeet = IsertSheet(MyWorkbook, "Prix Fils", True)
'
'DeleteRow MySeet, True
'
'Set MyRange = MySeet.Range("A5").CurrentRegion
''Myrange.Application.Visible = True
'For I = 0 To Rs.Fields.Count - 2
'    MyRange(1, I + 1) = Rs.Fields(I).Name
'Next
'Set MyRange = MySeet.Range("A5").CurrentRegion
'
'    MyRange.Interior.ColorIndex = 15
'    MyRange.HorizontalAlignment = xlCenter
'
'Row = 2
' FormBarGrah.ProgressBar1.Value = 0
' If NbLigne = 0 Then NbLigne = 1
' FormBarGrah.ProgressBar1.Max = NbLigne
' FormBarGrah.ProgressBar1Caption.Caption = " Exporter Prix du Câble :"
'While Rs.EOF = False
'     IncremanteBarGrah FormBarGrah
'    DoEvents
'    MyRange(Row, 1) = "" & Rs!TEINT
'    MyRange(Row, 2) = "" & Rs!Option
'    MyRange(Row, 3) = "" & Rs!ISO
'    MyRange(Row, 4) = Val(Replace("" & Rs!SECT, ",", "."))
'    MyRange(Row, 5) = Val(Replace("" & Rs!Longeur, ",", "."))
'    MyRange(Row, 6) = Val(Replace("" & Rs![Prix u], ",", "."))
'    MyRange(Row, 7).FormulaR1C1 = "" & Rs![Prix Total]
'
'    Rs.MoveNext
'    Row = Row + 1
'Wend
'Dim sql As String
'Set MyRange = MySeet.Range("A5").CurrentRegion
'
'sql = "SELECT  T_indiceProjet.Ensemble, T_indiceProjet.Equipement,[Li] & '_' & [LI_Indice] AS Liste, "
'sql = sql & "[PI] & '_' & [PI_Indice] AS Piece, T_indiceProjet.Client, T_indiceProjet.CleAc "
'sql = sql & "FROM T_indiceProjet "
'sql = sql & "WHERE T_indiceProjet.Id=" & Id_IndiceProjet & ";"
'Dim RsEntetePage As Recordset
'Set RsEntetePage = Con.OpenRecordSet(sql)
'
'MySeet.Range("F2") = "SOUS TOTAL"
'FormatExcelPlage MySeet.Range("F2"), 15, False, True, xlCenter, xlCenter
'R1 = MySeet.Range(MyRange(2, MyRange.Columns.Count).Address).Row
'R2 = MySeet.Range(MyRange(MyRange.Rows.Count, MyRange.Columns.Count).Address).Row
'MySeet.Range("G2").FormulaR1C1 = "=SUBTOTAL(9,R[" & R1 - 2 & "]C:R[" & R2 - 2 & "]C)"
'FormatExcelPlage MySeet.Range("G2"), 2, False, True, xlCenter, xlCenter
'
'MiseEnPage MySeet, MyRange, "Affaire : " & RsEntetePage!CleAc & Chr(10) & _
' _
'     "Pièce : " & RsEntetePage!Piece & Chr(10) & "Liste : " & RsEntetePage!Liste, vbCrLf & "Câblage : " & Replace("" & RsEntetePage!Ensemble, vbCrLf, " ") & Chr(10) & "Equipement : " & Replace("" & RsEntetePage!Equipement, vbCrLf, " ") & Chr(10) & _
'     "" _
'     , "ENCELADE" & Chr(10) & "Client : " & RsEntetePage!Client & Chr(10) & Format(Date, "dd-mmm-yyyy"), "", "&P/&N", "", 80, "B6", True, 1, True
'
'      MaJEncadreXls MyRange, xlThin, xlThin, xlHairline, xlHairline
'
'Set RsEntetePage = Con.CloseRecordSet(RsEntetePage)
'Set MyRange = Nothing
'insertExelAccess MySeet, "T_Prix_Fils", 5, Id_IndiceProjet
'Set MySeet = Nothing
'
'End Function
Public Sub insertExelAccess(MySheet As EXCEL.Worksheet, Table As String, RowStart As Long, Id_IndiceProjet As Long, _
                            Optional OnGletName As Boolean, Optional NotDeletTable As Boolean)
Dim I As Long
Dim I2 As Long
Dim Sql As String
Dim SqlValue As String
Dim MyRange As Range
Dim Rs As Recordset
Dim a
On Error GoTo 0
If NotDeletTable = False Then
    Sql = "DELETE " & Table & ".* FROM " & Table & " WHERE " & Table & ".Id_IndiceProjet=" & Id_IndiceProjet & ";"
    Con.Execute Sql
End If
Set Rs = Con.OpenRecordSet("SELECT " & Table & ".* FROM " & Table & " WHERE " & Table & ".ID=0;")

Set MyRange = MySheet.Cells(RowStart, 1).CurrentRegion
'Myrange.Application'.Visible = True
Sql = "INSERT INTO " & Table & " ( Id_IndiceProjet, "
If OnGletName = True Then Sql = Sql & "Onglet,"
For I = 1 To MyRange.Columns.Count
    Sql = Sql & "[" & MyRange(1, I) & "],"
Next
Sql = Left(Sql, Len(Sql) - 1) & ") Values (" & Id_IndiceProjet & ","
If OnGletName = True Then Sql = Sql & "'" & Replace(MySheet.Name, "'", "''") & "',"
If MyRange.Rows.Count = 1 Then

SqlValue = ""
        For I2 = 1 To MyRange.Columns.Count
            SqlValue = SqlValue & "null,"
        Next
        SqlValue = Left(SqlValue, Len(SqlValue) - 1) & ");"
    Con.Execute Sql & SqlValue

End If
For I = 2 To MyRange.Rows.Count
    SqlValue = ""
        For I2 = 1 To MyRange.Columns.Count
'        Debug.Print Myrange(I, I2).Address
'       Debug.Print Myrange(1, I2).Value & " = " & MySheet.Range(Myrange(I, I2).Address).FormulaR1C1
'       Myrange.Application'.Visible = True

Debug.Print MyRange(1, I2).Value & " : " & MyRange(1, I2).Value; a; " " & "" & MyRange(I, I2).FormulaR1C1
'Myrange.Application '.Visible = True
        Select Case Rs(MyRange(1, I2).Value).Type
        Case 11
            SqlValue = SqlValue & Replace(Replace(Replace(Replace(UCase(MyRange(I, I2)), "FALSE", 0), "TRUE", 1), "FAUX", 0), "VRAI", 1) & ","
        Case 202
            SqlValue = SqlValue & "'" & MyReplace("" & MyRange(I, I2).FormulaR1C1) & "',"
        Case 203
            SqlValue = SqlValue & "'" & MyReplace("" & MyRange(I, I2).FormulaR1C1) & "',"
        Case 5
            SqlValue = SqlValue & Replace(Val(Replace("" & MyReplace(MyRange(I, I2).FormulaR1C1), ",", ".")), ",", ".") & ","
        Case 3
            SqlValue = SqlValue & Replace(Val(Replace("" & MyReplace(MyRange(I, I2).FormulaR1C1), ",", ".")), ",", ".") & ","
        Case Else
            MsgBox ""
        End Select
    Next
    SqlValue = Left(SqlValue, Len(SqlValue) - 1) & ");"
    Con.Execute Sql & SqlValue
    
    
Next

Set Rs = Con.CloseRecordSet(Rs)
End Sub

Sub MaJEncadreXls(MyRange As Range, LeftWeight As Long, RightWeight As Long, TopWeight As Long, BottomWeight As Long)
On Error Resume Next

'
' Macro3 Macro
' Macro enregistrée le 14/03/2005 par robert.durupt
'

'
    MyRange.Borders(xlDiagonalDown).LineStyle = xlNone
    MyRange.Borders(xlDiagonalUp).LineStyle = xlNone
    
    MyRange.Borders(xlEdgeLeft).LineStyle = xlContinuous
    MyRange.Borders(xlEdgeLeft).Weight = LeftWeight
   
        MyRange.Borders(xlEdgeRight).LineStyle = xlContinuous
        MyRange.Borders(xlEdgeRight).Weight = RightWeight
        
      
        MyRange.Borders(xlEdgeTop).LineStyle = xlContinuous
        MyRange.Borders(xlEdgeTop).Weight = TopWeight
       
    
        MyRange.Borders(xlEdgeBottom).LineStyle = xlContinuous
        MyRange.Borders(xlEdgeBottom).Weight = BottomWeight
      
    MyRange.Borders(xlInsideVertical).Weight = LeftWeight
     MyRange.Borders(xlInsideHorizontal).Weight = BottomWeight
 On Error GoTo 0
End Sub


Sub FormatExcelPlage(Plage As Range, Couleur As Long, Merge As Boolean, Grille As Boolean, HorizontalAlignment As Long, VerticalAlignment As Long)
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
Sub DeletSheet(MySheet As EXCEL.Worksheet)
    On Error Resume Next
    MySheet.Delete
Err.Clear
End Sub
Sub MajBase(IdIndice As Long)
Dim Sql As String
'***********************************************************************************************************************
'*                                        Supprime les données des tables de travail :                                 *
Sql = "DELETE T_Critères.* "
Sql = Sql & "FROM T_Critères "
Sql = Sql & "WHERE T_Critères.Id_IndiceProjet=" & IdIndice & ";"
Con.Execute Sql

Sql = "DELETE Ligne_Tableau_fils.* "
Sql = Sql & "FROM Ligne_Tableau_fils "
Sql = Sql & "WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & IdIndice & ";"
Con.Execute Sql

Sql = "DELETE Connecteurs.* "
Sql = Sql & "FROM Connecteurs "
Sql = Sql & "WHERE Connecteurs.Id_IndiceProjet=" & IdIndice & ";"
Con.Execute Sql

Sql = "DELETE Composants.* "
Sql = Sql & "FROM Composants "
Sql = Sql & "WHERE Composants.Id_IndiceProjet=" & IdIndice & ";"
Con.Execute Sql

Sql = "DELETE Nota.* "
Sql = Sql & "FROM Nota "
Sql = Sql & "WHERE Nota.Id_IndiceProjet=" & IdIndice & ";"
Con.Execute Sql


Sql = "DELETE T_Noeuds.* "
Sql = Sql & "FROM T_Noeuds "
Sql = Sql & "WHERE T_Noeuds.Id_IndiceProjet=" & IdIndice & ";"
Con.Execute Sql
'***********************************************************************************************************************
'*                                        Enrichie les données des tables de travail :                                 *


Sql = "INSERT INTO T_Critères ( Id_IndiceProjet,ACTIVER,CODE_CRITERE, CRITERES,COMMENTAIRES)  "
Sql = Sql & "SELECT " & IdIndice & " AS Id_IndiceProjet,Xls_Critères.ACTIVER, Xls_Critères.CODE_CRITERE, Xls_Critères.CRITERES ,Xls_Critères.COMMENTAIRES "
Sql = Sql & "FROM Xls_Critères  "
Sql = Sql & "WHERE Xls_Critères.Job=" & NmJob & ";"
Con.Execute Sql
'
'Sql = "INSERT INTO Ligne_Tableau_fils ( Id_IndiceProjet,[_Client2],ACTIVER,LIAI, DESIGNATION,  "
'Sql = Sql & "FIL, SECT, TEINT, TEINT2, ISO, [LONG], [LONG CP], COUPE, POS,  "
'Sql = Sql & "[POS-OUT], FA, APP, VOI, POS2, [POS-OUT2], FA2, APP2,  "
'Sql = Sql & "VOI2, PRECO, [OPTION],[Critères spécifiques] ) "
'Sql = Sql & "SELECT " & IdIndice & " AS Id_IndiceProjet, "
'Sql = Sql & " "
'Sql = Sql & " "
'Sql = Sql & ", "
'Sql = Sql & "xls_Ligne_Tableau_fils.ACTIVER,  "
'Sql = Sql & "xls_Ligne_Tableau_fils.LIAI,  "
'Sql = Sql & " xls_Ligne_Tableau_fils.DESIGNATION, xls_Ligne_Tableau_fils.FIL,  "
'Sql = Sql & "xls_Ligne_Tableau_fils.SECT, xls_Ligne_Tableau_fils.TEINT,  "
'Sql = Sql & "xls_Ligne_Tableau_fils.TEINT2, xls_Ligne_Tableau_fils.ISO,  "
'Sql = Sql & "xls_Ligne_Tableau_fils.LONG, xls_Ligne_Tableau_fils.[LONG CP],  "
'Sql = Sql & "xls_Ligne_Tableau_fils.COUPE, xls_Ligne_Tableau_fils.POS,  "
'Sql = Sql & "xls_Ligne_Tableau_fils.[POS-OUT], xls_Ligne_Tableau_fils.FA,  "
'Sql = Sql & "xls_Ligne_Tableau_fils.APP, xls_Ligne_Tableau_fils.VOI,  "
'Sql = Sql & "xls_Ligne_Tableau_fils.POS2, xls_Ligne_Tableau_fils.[POS-OUT2],  "
'Sql = Sql & "xls_Ligne_Tableau_fils.FA2, xls_Ligne_Tableau_fils.APP2,  "
'Sql = Sql & "xls_Ligne_Tableau_fils.VOI2, xls_Ligne_Tableau_fils.PRECO,  "
'Sql = Sql & "xls_Ligne_Tableau_fils.OPTION ,xls_Ligne_Tableau_fils.[Critères spécifiques] "
'Sql = Sql & "FROM xls_Ligne_Tableau_fils "
'Sql = Sql & " where xls_Ligne_Tableau_fils.Job=" & NmJob & ";"

'sql = "INSERT INTO Ligne_Tableau_fils ( Id_IndiceProjet, LIAI, DESIGNATION, FIL, SECT,  "
'sql = sql & "TEINT, TEINT2, ISO, [LONG], [LONG CP], COUPE, POS, [POS-OUT], FA, APP, VOI,  "
'sql = sql & "[Ref Connecteur],[Ref Connecteur_Four],[Long_Add], [Ref Clip], [Ref Clip Four], [Ref Joint], [Ref Joint Four],  "
'sql = sql & "POS2, [POS-OUT2], FA2, APP2, VOI2,[Ref Connecteur2],[Ref Connecteur_Four2],[Long_Add2], [Ref Clip2], [Ref Clip Four2],  "
'sql = sql & "[Ref Joint2], [Ref Joint Four2], PRECOG, [OPTION], ACTIVER, [Critères spécifiques] ) "
'sql = sql & "SELECT " & IdIndice & " AS Id_IndiceProjet, xls_Ligne_Tableau_fils.LIAI,  "
'sql = sql & "xls_Ligne_Tableau_fils.DESIGNATION,xls_Ligne_Tableau_fils.FIL, xls_Ligne_Tableau_fils.SECT,  "
'sql = sql & "xls_Ligne_Tableau_fils.TEINT, xls_Ligne_Tableau_fils.TEINT2, xls_Ligne_Tableau_fils.ISO,  "
'sql = sql & "xls_Ligne_Tableau_fils.LONG, xls_Ligne_Tableau_fils.[LONG CP], xls_Ligne_Tableau_fils.COUPE,  "
'sql = sql & "xls_Ligne_Tableau_fils.POS, xls_Ligne_Tableau_fils.[POS-OUT], xls_Ligne_Tableau_fils.FA,  "
'sql = sql & "xls_Ligne_Tableau_fils.APP, xls_Ligne_Tableau_fils.VOI,  "
'sql = sql & "xls_Ligne_Tableau_fils.[Ref Connecteur],xls_Ligne_Tableau_fils.[Ref Connecteur_Four],xls_Ligne_Tableau_fils.[Long_Add], xls_Ligne_Tableau_fils.[Ref Clip],  "
'sql = sql & "xls_Ligne_Tableau_fils.[Ref Clip Four], xls_Ligne_Tableau_fils.[Ref Joint],  "
'sql = sql & "xls_Ligne_Tableau_fils.[Ref Joint Four], xls_Ligne_Tableau_fils.POS2,  "
'sql = sql & "xls_Ligne_Tableau_fils.[POS-OUT2], xls_Ligne_Tableau_fils.FA2,  "
'sql = sql & "xls_Ligne_Tableau_fils.APP2, xls_Ligne_Tableau_fils.VOI2,  "
'sql = sql & "xls_Ligne_Tableau_fils.[Ref Connecteur2],xls_Ligne_Tableau_fils.[Ref Connecteur_Four2],xls_Ligne_Tableau_fils.[Long_Add2], xls_Ligne_Tableau_fils.[Ref Clip2],  "
'sql = sql & "xls_Ligne_Tableau_fils.[Ref Clip Four2], xls_Ligne_Tableau_fils.[Ref Joint2],  "
'sql = sql & "xls_Ligne_Tableau_fils.[Ref Joint Four2], xls_Ligne_Tableau_fils.PRECOG ,  "
'sql = sql & "xls_Ligne_Tableau_fils.Option, xls_Ligne_Tableau_fils.Activer,  "
'sql = sql & "xls_Ligne_Tableau_fils.[Critères spécifiques] "
'sql = sql & "FROM xls_Ligne_Tableau_fils "
'Sql = Sql & "WHERE xls_Ligne_Tableau_fils.Job=" & NmJob & ";"
'
'Sql = "INSERT INTO Ligne_Tableau_fils ( Id_IndiceProjet, LIAI, DESIGNATION, FIL, SECT, TEINT,   "
'Sql = Sql & "TEINT2, ISO, [LONG], [LONG CP], COUPE, POS, [POS-OUT], FA, APP, VOI, [Ref Connecteur],[Ref Connecteur_Four],[Long_Add],   "
'Sql = Sql & "[Ref Clip], [Ref Clip Four], [Ref Joint], [Ref Joint Four], POS2, [POS-OUT2],   "
'Sql = Sql & "FA2, APP2, VOI2, [Ref Connecteur2],[Ref Connecteur_Four2],[Long_Add2], [Ref Clip2], [Ref Clip Four2], [Ref Joint2],   "
'Sql = Sql & "[Ref Joint Four2], PRECOG, [OPTION], ACTIVER, [Critères spécifiques]   "
'Sql = Sql & " )  "
'Sql = Sql & "SELECT " & IdIndice & " AS Id_IndiceProjet, xls_Ligne_Tableau_fils.LIAI,   "
'Sql = Sql & "xls_Ligne_Tableau_fils.DESIGNATION,xls_Ligne_Tableau_fils.FIL,   "
'Sql = Sql & "xls_Ligne_Tableau_fils.SECT, xls_Ligne_Tableau_fils.TEINT,   "
'Sql = Sql & "xls_Ligne_Tableau_fils.TEINT2, xls_Ligne_Tableau_fils.ISO,   "
'Sql = Sql & "xls_Ligne_Tableau_fils.LONG, xls_Ligne_Tableau_fils.[LONG CP],   "
'Sql = Sql & "xls_Ligne_Tableau_fils.COUPE, xls_Ligne_Tableau_fils.POS, xls_Ligne_Tableau_fils.[POS-OUT],   "
'Sql = Sql & "xls_Ligne_Tableau_fils.FA, xls_Ligne_Tableau_fils.APP, xls_Ligne_Tableau_fils.VOI,   "
'Sql = Sql & "xls_Ligne_Tableau_fils.[Ref Connecteur],xls_Ligne_Tableau_fils.[Ref Connecteur_Four],xls_Ligne_Tableau_fils.[Long_Add], xls_Ligne_Tableau_fils.[Ref Clip],   "
'Sql = Sql & "xls_Ligne_Tableau_fils.[Ref Clip Four], xls_Ligne_Tableau_fils.[Ref Joint],   "
'Sql = Sql & "xls_Ligne_Tableau_fils.[Ref Joint Four], xls_Ligne_Tableau_fils.POS2,   "
'Sql = Sql & "xls_Ligne_Tableau_fils.[POS-OUT2], xls_Ligne_Tableau_fils.FA2,   "
'Sql = Sql & "xls_Ligne_Tableau_fils.APP2, xls_Ligne_Tableau_fils.VOI2,   "
'Sql = Sql & "xls_Ligne_Tableau_fils.[Ref Connecteur2],xls_Ligne_Tableau_fils.[Ref Connecteur_Four2],xls_Ligne_Tableau_fils.[Long_Add2], xls_Ligne_Tableau_fils.[Ref Clip2],   "
'Sql = Sql & "xls_Ligne_Tableau_fils.[Ref Clip Four2], xls_Ligne_Tableau_fils.[Ref Joint2],   "
'Sql = Sql & "xls_Ligne_Tableau_fils.[Ref Joint Four2], xls_Ligne_Tableau_fils.PRECOG ,   "
'Sql = Sql & "xls_Ligne_Tableau_fils.Option, xls_Ligne_Tableau_fils.ACTIVER,   "
'Sql = Sql & "xls_Ligne_Tableau_fils.[Critères spécifiques] "
'Sql = Sql & "FROM xls_Ligne_Tableau_fils  "
'Sql = Sql & "WHERE xls_Ligne_Tableau_fils.Job=" & NmJob & ";"
'

Sql = "INSERT INTO Ligne_Tableau_fils ( Id_IndiceProjet, LIAI, DESIGNATION, FIL, SECT, TEINT, TEINT2,  "
Sql = Sql & "ISO, [LONG], [LONG CP], COUPE, POS, [POS-OUT], FA, APP, VOI, [Ref Connecteur], [Ref Connecteur_Four],  "
Sql = Sql & "Long_Add, [Ref Clip], [Ref Clip Four], [Ref Joint], [Ref Joint Four], POS2, [POS-OUT2], FA2, APP2, VOI2,  "
Sql = Sql & "[Ref Connecteur2], [Ref Connecteur_Four2], Long_Add2, [Ref Clip2], [Ref Clip Four2], [Ref Joint2],  "
Sql = Sql & "[Ref Joint Four2], PRECOG, [OPTION], ACTIVER, [Critères spécifiques], PRECO,PRECO1,COMMENTAIRES) "
Sql = Sql & "SELECT  " & IdIndice & " AS Id_IndiceProjet, xls_Ligne_Tableau_fils.LIAI, xls_Ligne_Tableau_fils.DESIGNATION,  "
Sql = Sql & "xls_Ligne_Tableau_fils.FIL, xls_Ligne_Tableau_fils.SECT, xls_Ligne_Tableau_fils.TEINT,  "
Sql = Sql & "xls_Ligne_Tableau_fils.TEINT2, xls_Ligne_Tableau_fils.ISO, xls_Ligne_Tableau_fils.LONG,  "
Sql = Sql & "xls_Ligne_Tableau_fils.[LONG CP], xls_Ligne_Tableau_fils.COUPE, xls_Ligne_Tableau_fils.POS,  "
Sql = Sql & "xls_Ligne_Tableau_fils.[POS-OUT], xls_Ligne_Tableau_fils.FA, xls_Ligne_Tableau_fils.APP,  "
Sql = Sql & "xls_Ligne_Tableau_fils.VOI, xls_Ligne_Tableau_fils.[Ref Connecteur], xls_Ligne_Tableau_fils.[Ref Connecteur_Four],  "
Sql = Sql & "xls_Ligne_Tableau_fils.Long_Add, xls_Ligne_Tableau_fils.[Ref Clip], xls_Ligne_Tableau_fils.[Ref Clip Four],  "
Sql = Sql & "xls_Ligne_Tableau_fils.[Ref Joint], xls_Ligne_Tableau_fils.[Ref Joint four], xls_Ligne_Tableau_fils.POS2,  "
Sql = Sql & "xls_Ligne_Tableau_fils.[POS-OUT2], xls_Ligne_Tableau_fils.FA2, xls_Ligne_Tableau_fils.APP2, xls_Ligne_Tableau_fils.VOI2,  "
Sql = Sql & "xls_Ligne_Tableau_fils.[Ref Connecteur2], xls_Ligne_Tableau_fils.[Ref Connecteur_Four2], xls_Ligne_Tableau_fils.Long_Add2 "
Sql = Sql & ", xls_Ligne_Tableau_fils.[Ref Clip2], xls_Ligne_Tableau_fils.[Ref Clip Four2], xls_Ligne_Tableau_fils.[Ref Joint2],  "
Sql = Sql & "xls_Ligne_Tableau_fils.[Ref Joint four2], xls_Ligne_Tableau_fils.PRECOG, xls_Ligne_Tableau_fils.OPTION,  "
Sql = Sql & "xls_Ligne_Tableau_fils.ACTIVER, xls_Ligne_Tableau_fils.[Critères spécifiques], xls_Ligne_Tableau_fils.PRECO,  "
Sql = Sql & "xls_Ligne_Tableau_fils.PRECO1 "
Sql = Sql & "FROM xls_Ligne_Tableau_fils "
Sql = Sql & "WHERE xls_Ligne_Tableau_fils.Job=" & NmJob & ";"



Sql = "INSERT INTO Ligne_Tableau_fils ( Id_IndiceProjet, LIAI, DESIGNATION, FIL, SECT, TEINT, TEINT2,  "
Sql = Sql & "ISO, [LONG], [LONG CP], COUPE, POS, [POS-OUT], FA, APP, VOI, [Ref Connecteur], [Ref Connecteur_Four],  "
Sql = Sql & "Long_Add, [Ref Clip], [Ref Clip Four], [Ref Joint], [Ref Joint Four], POS2, [POS-OUT2], FA2, APP2, VOI2,  "
Sql = Sql & "[Ref Connecteur2], [Ref Connecteur_Four2], Long_Add2, [Ref Clip2], [Ref Clip Four2], [Ref Joint2],  "
Sql = Sql & "[Ref Joint Four2], PRECOG, [OPTION], ACTIVER, [Critères spécifiques], PRECO, PRECO2,COMMENTAIRES  ) "
Sql = Sql & "SELECT " & IdIndice & " AS Id_IndiceProjet, xls_Ligne_Tableau_fils.LIAI, xls_Ligne_Tableau_fils.DESIGNATION, xls_Ligne_Tableau_fils.FIL,  "
Sql = Sql & "xls_Ligne_Tableau_fils.SECT, xls_Ligne_Tableau_fils.TEINT, xls_Ligne_Tableau_fils.TEINT2, xls_Ligne_Tableau_fils.ISO,  "
Sql = Sql & "xls_Ligne_Tableau_fils.LONG, xls_Ligne_Tableau_fils.[LONG CP], xls_Ligne_Tableau_fils.COUPE, xls_Ligne_Tableau_fils.POS,  "
Sql = Sql & "xls_Ligne_Tableau_fils.[POS-OUT], xls_Ligne_Tableau_fils.FA, xls_Ligne_Tableau_fils.APP, xls_Ligne_Tableau_fils.VOI,  "
Sql = Sql & "xls_Ligne_Tableau_fils.[Ref Connecteur], xls_Ligne_Tableau_fils.[Ref Connecteur_Four], xls_Ligne_Tableau_fils.Long_Add,  "
Sql = Sql & "xls_Ligne_Tableau_fils.[Ref Clip], xls_Ligne_Tableau_fils.[Ref Clip Four], xls_Ligne_Tableau_fils.[Ref Joint],  "
Sql = Sql & "xls_Ligne_Tableau_fils.[Ref Joint four], xls_Ligne_Tableau_fils.POS2, xls_Ligne_Tableau_fils.[POS-OUT2], xls_Ligne_Tableau_fils.FA2,  "
Sql = Sql & "xls_Ligne_Tableau_fils.APP2, xls_Ligne_Tableau_fils.VOI2, xls_Ligne_Tableau_fils.[Ref Connecteur2],  "
Sql = Sql & "xls_Ligne_Tableau_fils.[Ref Connecteur_Four2], xls_Ligne_Tableau_fils.Long_Add2,  "
Sql = Sql & "xls_Ligne_Tableau_fils.[Ref Clip2], xls_Ligne_Tableau_fils.[Ref Clip Four2], xls_Ligne_Tableau_fils.[Ref Joint2],  "
Sql = Sql & "xls_Ligne_Tableau_fils.[Ref Joint four2], xls_Ligne_Tableau_fils.PRECOG, xls_Ligne_Tableau_fils.OPTION,  "
Sql = Sql & "xls_Ligne_Tableau_fils.ACTIVER, xls_Ligne_Tableau_fils.[Critères spécifiques], xls_Ligne_Tableau_fils.PRECO,  "
Sql = Sql & " xls_Ligne_Tableau_fils.PRECO2,xls_Ligne_Tableau_fils.COMMENTAIRES "
Sql = Sql & "FROM xls_Ligne_Tableau_fils "
Sql = Sql & "WHERE xls_Ligne_Tableau_fils.Job=" & NmJob & ";"





Con.Execute Sql
'Sql = "INSERT INTO Connecteurs ( Id_IndiceProjet, ACTIVER, CONNECTEUR, RefConnecteurFour, [O/N],  "
'Sql = Sql & "DESIGNATION, CODE_APP, N°, POS, [POS-OUT], PRECO1, PRECO2, [100%], [OPTION], Pylone,  "
'Sql = Sql & "Colonne, Ligne, RefBouchon, ReFCapot, RefVerrou,LongueurF_Choix ) "
'Sql = Sql & "SELECT " & IdIndice & " AS Id_IndiceProjet, Xls_Connecteurs.ACTIVER, Xls_Connecteurs.CONNECTEUR,  "
'Sql = Sql & "Xls_Connecteurs.RefConnecteurFour, Xls_Connecteurs.[O/N], Xls_Connecteurs.DESIGNATION,  "
'Sql = Sql & "Xls_Connecteurs.CODE_APP, Xls_Connecteurs.N°, Xls_Connecteurs.POS,  "
'Sql = Sql & "Xls_Connecteurs.[POS-OUT], Xls_Connecteurs.PRECO1, Xls_Connecteurs.PRECO2,  "
'Sql = Sql & "Xls_Connecteurs.[100%], Xls_Connecteurs.OPTION, Xls_Connecteurs.Pylone,  "
'Sql = Sql & "Xls_Connecteurs.Colonne, Xls_Connecteurs.Ligne, Xls_Connecteurs.RefBouchon,  "
'Sql = Sql & "Xls_Connecteurs.ReFCapot, Xls_Connecteurs.RefVerrou,Xls_Connecteurs.LongueurF_Choix "
'Sql = Sql & "FROM Xls_Connecteurs "
'Sql = Sql & "WHERE Xls_Connecteurs.Job=" & NmJob & ";"

Sql = "INSERT INTO Connecteurs ( Id_IndiceProjet, ACTIVER, CONNECTEUR, RefConnecteurFour, [O/N], DESIGNATION, CODE_APP, N°, "
Sql = Sql & "POS, [POS-OUT], PRECO1, PRECO2, [100%], [OPTION], Pylone, Colonne, Ligne, RefBouchon, RefBouchonFour, ReFCapot, "
Sql = Sql & "ReFCapotFour, RefVerrou, RefVerrouFour, LongueurF_Choix,COMMENTAIRES  )"
Sql = Sql & "SELECT " & IdIndice & " AS Id_IndiceProjet, Xls_Connecteurs.ACTIVER, Xls_Connecteurs.CONNECTEUR, "
Sql = Sql & "Xls_Connecteurs.RefConnecteurFour, Xls_Connecteurs.[O/N], Xls_Connecteurs.DESIGNATION, "
Sql = Sql & "Xls_Connecteurs.CODE_APP, Xls_Connecteurs.N°, Xls_Connecteurs.POS, Xls_Connecteurs.[POS-OUT], "
Sql = Sql & "Xls_Connecteurs.PRECO1,Xls_Connecteurs.PRECO2, Xls_Connecteurs.[100%], Xls_Connecteurs.OPTION, Xls_Connecteurs.Pylone, "
Sql = Sql & "Xls_Connecteurs.Colonne, Xls_Connecteurs.Ligne, Xls_Connecteurs.RefBouchon, Xls_Connecteurs.RefBouchonFour, "
Sql = Sql & "Xls_Connecteurs.ReFCapot, Xls_Connecteurs.ReFCapotFour, Xls_Connecteurs.RefVerrou, Xls_Connecteurs.RefVerrouFour, "
Sql = Sql & "Xls_Connecteurs.LongueurF_Choix,Xls_Connecteurs.COMMENTAIRES  "
Sql = Sql & "FROM Xls_Connecteurs "
Sql = Sql & "WHERE Xls_Connecteurs.Job=" & NmJob & ";"


Con.Execute Sql



Sql = "INSERT INTO Composants ( Id_IndiceProjet, ACTIVER,DESIGNCOMP, NUMCOMP, REFCOMP, Path  ,[OPTION],Code_APP_Lier,Voie,POS,[POS-OUT] ,COMMENTAIRES ) "
Sql = Sql & "SELECT " & IdIndice & " AS Id_IndiceProjet, Xls_Composants.ACTIVER,Xls_Composants.DESIGNCOMP, Xls_Composants.NUMCOMP,   "
Sql = Sql & "Xls_Composants.REFCOMP, Xls_Composants.Path  , Xls_Composants.[OPTION] ,Xls_Composants.Code_APP_Lier,Xls_Composants.Voie,Xls_Composants.POS,Xls_Composants.[POS-OUT] ,Xls_Composants.COMMENTAIRES  "
Sql = Sql & "FROM Xls_Composants "
Sql = Sql & "WHERE Xls_Composants.Job=" & NmJob & ";"

Con.Execute Sql

Sql = "INSERT INTO Nota ( Id_IndiceProjet,ACTIVER, NOTA, NUMNOTA,[OPTION],COMMENTAIRES  ) "
Sql = Sql & "SELECT " & IdIndice & " AS Id_IndiceProjet, Xls_Nota.ACTIVER,Xls_Nota.NOTA, Xls_Nota.NUMNOTA,Xls_Nota.[OPTION] ,Xls_Nota.COMMENTAIRES "
Sql = Sql & "FROM Xls_Nota "
Sql = Sql & "WHERE Xls_Nota.Job=" & NmJob & ";"

Con.Execute Sql

Sql = "INSERT INTO T_Noeuds ( Id_IndiceProjet,Fleche_Droite, ACTIVER, NŒUDS,LONGUEUR,DESIGN_HAB, "
Sql = Sql & "CODE_RSA,CODE_PSA,CODE_ENC,DIAMETRE,CLASSE_T,TORON_PRINCIPAL, LONGUEUR_CUMULEE,[OPTION],COMMENTAIRES) "
Sql = Sql & "SELECT " & IdIndice & " AS Id_IndiceProjet, Xls_Noeuds.Fleche_Droite, Xls_Noeuds.ACTIVER, "
Sql = Sql & "Xls_Noeuds.NŒUDS,Xls_Noeuds.LONGUEUR,Xls_Noeuds.DESIGN_HAB,Xls_Noeuds.CODE_RSA, "
Sql = Sql & "Xls_Noeuds.CODE_PSA,Xls_Noeuds.CODE_ENC,Xls_Noeuds.DIAMETRE,Xls_Noeuds.CLASSE_T,Xls_Noeuds.TORON_PRINCIPAL, "
Sql = Sql & "Xls_Noeuds.LONGUEUR_CUMULEE ,Xls_Noeuds.[OPTION],Xls_Noeuds.COMMENTAIRES "
Sql = Sql & "FROM Xls_Noeuds "
Sql = Sql & "WHERE Xls_Noeuds.Job=" & NmJob & ";"

Con.Execute Sql


Sql = "DELETE Xls_Critères.*  FROM Xls_Critères "
Sql = Sql & "where Xls_Critères.Job=" & NmJob & ";"
Con.Execute Sql


 Sql = "DELETE xls_Ligne_Tableau_fils.*  FROM xls_Ligne_Tableau_fils "
Sql = Sql & "where xls_Ligne_Tableau_fils.Job=" & NmJob & ";"
Con.Execute Sql

Sql = "DELETE Xls_Composants.*  FROM Xls_Composants "
Sql = Sql & "where Xls_Composants.Job=" & NmJob & ";"
Con.Execute Sql

Sql = "DELETE Xls_Nota.*  FROM Xls_Nota "
Sql = Sql & "where Xls_Nota.Job=" & NmJob & ";"
Con.Execute Sql

Sql = "DELETE Xls_Connecteurs.* FROM Xls_Connecteurs "
Sql = Sql & "WHERE Xls_Connecteurs.Job=" & NmJob & ";"
Con.Execute Sql

'***********************************************************************************************************************
'*                                        Attribut les code appareil au tableau de fils :                              *

Sql = "UPDATE (Ligne_Tableau_fils LEFT JOIN Connecteurs ON Ligne_Tableau_fils.FA = Connecteurs.N°)  "
Sql = Sql & "LEFT JOIN Connecteurs AS Connecteurs_1 ON Ligne_Tableau_fils.FA2 = Connecteurs_1.N°  "
Sql = Sql & "SET Ligne_Tableau_fils.APP = [Connecteurs].[CODE_APP], Ligne_Tableau_fils.APP2 = [Connecteurs_1].[CODE_APP] "
Sql = Sql & "WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & IdIndice & " "
Sql = Sql & "AND Connecteurs.Id_IndiceProjet=" & IdIndice & " "
Sql = Sql & "AND Connecteurs_1.Id_IndiceProjet=" & IdIndice & ";"
Con.Execute Sql
'***********************************************************************************************************************

End Sub

Sub Racourci(RaccourciName As String, RaccourciCible As String, extension As String)
Dim Fso As New FileSystemObject
Dim objshell, objraccourci
If Fso.FileExists(RaccourciName & ".Lnk") = True Then
     Fso.DeleteFile RaccourciName & ".Lnk"
End If
Set objshell = CreateObject("wscript.shell")
Set objraccourci = objshell.createshortcut(RaccourciName & ".Lnk")
objraccourci.targetpath = RaccourciCible & "." & extension
objraccourci.Save
Set Fso = Nothing
Set objraccourci = Nothing
End Sub
'Function ExporteXlsHabillages(Rs As Recordset, Id_IndiceProjet As Long)
'Dim MySeet As Excel.Worksheet
'Dim MyRange As Excel.Range
''Dim MyWorkbook As Workbook
'Dim Row As Long
'Dim NbLigne As Long
'Dim I As Long
'Dim R1 As Long
'Dim R2 As Long
'DoEvents
'    NbLigne = 0
'If Rs.EOF = True Then Exit Function
'While Rs.EOF = False
'NbLigne = NbLigne + 1
'Rs.MoveNext
'Wend
'Rs.Requery
'
'Set MySeet = IsertSheet(MyWorkbook, "Appro_Habillage", True)
''MySeet.Application.Visible = True
'DeleteRow MySeet, True
'
'Set MyRange = MySeet.Range("A1").CurrentRegion
''Myrange.Application.Visible = True
'Row = 6
' FormBarGrah.ProgressBar1.Value = 0
' If NbLigne = 0 Then NbLigne = 1
' FormBarGrah.ProgressBar1.Max = NbLigne
' FormBarGrah.ProgressBar1Caption.Caption = " Exporter liste des Habilage :"
' For I = 0 To Rs.Fields.Count - 2
'    MySeet.Cells(5, I + 1) = Rs.Fields(I).Name
'
' Next
'While Rs.EOF = False
'     IncremanteBarGrah FormBarGrah
'     For I = 0 To Rs.Fields.Count - 2
'
'
'         If Rs(I).Name = "Prix Total" Then
'            MySeet.Cells(Row, I + 1).FormulaR1C1 = "=(RC[-1]*RC[-2])"
'         End If
' Next
'    DoEvents
'    For I = 0 To Rs.Fields.Count - 2
'      If Rs(I).Name <> "Prix Total" Then
'
'
'        MySeet.Cells(Row, I + 1) = Trim(Replace("" & Rs(I), vbCrLf, ""))
'    End If
'    Next
'    Row = Row + 1
'    Rs.MoveNext
'Wend
'Dim sql As String
'Set MyRange = MySeet.Range("A5").CurrentRegion
'
'sql = "SELECT  T_indiceProjet.Ensemble, T_indiceProjet.Equipement,[Li] & '_' & [LI_Indice] AS Liste, "
'sql = sql & "[PI] & '_' & [PI_Indice] AS Piece, T_indiceProjet.Client, T_indiceProjet.CleAc "
'sql = sql & "FROM T_indiceProjet "
'sql = sql & "WHERE T_indiceProjet.Id=" & Id_IndiceProjet & ";"
'Dim RsEntetePage As Recordset
'Set RsEntetePage = Con.OpenRecordSet(sql)
'
'MySeet.Range("D2") = "SOUS TOTAL"
'FormatExcelPlage MySeet.Range("D2"), 15, False, True, xlCenter, xlCenter
'R1 = MySeet.Range(MyRange(2, MyRange.Columns.Count).Address).Row
'R2 = MySeet.Range(MyRange(MyRange.Rows.Count, MyRange.Columns.Count).Address).Row
'MySeet.Range("E2").FormulaR1C1 = "=SUBTOTAL(9,R[" & R1 - 2 & "]C:R[" & R2 - 2 & "]C)"
'FormatExcelPlage MySeet.Range("E2"), 2, False, True, xlCenter, xlCenter
'
'MiseEnPage MySeet, MyRange, "Affaire : " & RsEntetePage!CleAc & Chr(10) & _
' _
'     "Pièce : " & RsEntetePage!Piece & Chr(10) & "Liste : " & RsEntetePage!Liste, vbCrLf & "Câblage : " & Replace("" & RsEntetePage!Ensemble, vbCrLf, " ") & Chr(10) & "Equipement : " & Replace("" & RsEntetePage!Equipement, vbCrLf, " ") & Chr(10) & _
'     "" _
'    , "ENCELADE" & Chr(10) & "Client : " & RsEntetePage!Client & Chr(10) & Format(Date, "dd-mmm-yyyy"), "", "&P/&N", "", 80, "B6", True, 1, True
'
'      MaJEncadreXls MyRange, xlThin, xlThin, xlHairline, xlHairline
'
'Set RsEntetePage = Con.CloseRecordSet(RsEntetePage)
'Set MyRange = Nothing
'insertExelAccess MySeet, "T_Appro_Habillage", 5, Id_IndiceProjet
'
'Set MySeet = Nothing
'
'
'End Function
Public Sub ExcelCreatTitre(MyRangeStrate As Object, RsTitre As Recordset, Optional Value As Boolean, Optional ChampEnTire As Boolean, Optional Formule As Boolean)
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

Public Sub RazFiltreEditExcel(MySpreadsheet As Object)
On Error Resume Next
 MySpreadsheet.ActiveSheet.AutoFilterMode = False
'If MySpreadsheet.ActiveSheet.AutoFilterMode = True Then
'    For I = 1 To Myrange.Columns.Count
'    Set aa = MySpreadsheet.ActiveSheet.AutoFilter.Filters(I).Criteria
'    aa.Show All = True
'
'    Next
   MySpreadsheet.ActiveSheet.AutoFilterMode = True
    MySpreadsheet.ActiveSheet.Range("A1").AutoFilter
    MySpreadsheet.ActiveSheet.AutoFilter.Apply
DoEvents
'End If

End Sub

Public Sub Copy_Rs_Spreadsheet(FRM As Form, Spreadsheet, Rs As Recordset, Mytype As String, FrmApelan As Object, LibAction As String)
FrmApelan.ProgressBar1Caption = LibAction
DoEvents
Dim toto
On Error Resume Next
Dim L As Long
Dim I As Long
Dim Save_ProgressBar1Value As Long
Dim Save_ProgressBar1Max As Long
Save_ProgressBar1Value = FrmApelan.ProgressBar1.Value
Save_ProgressBar1Max = FrmApelan.ProgressBar1.Value
For I = 0 To Rs.Fields.Count - 1
DoEvents
    Spreadsheet.Cells(1, I + 1) = "'" & Rs(I).Name
    Spreadsheet.Cells(1, I + 1).Interior.Color = ChoixCouleur(0)
Next
Const sDelimiteur$ = vbTab
If Rs.EOF = False Then
    toto = Rs.GetString(, , sDelimiteur$ & "'", "¤")
     toto = Replace(toto, Chr(10), "©")
      toto = Replace(toto, Chr(13), "")
    toto = Replace(toto, "¤", vbCrLf)
    Spreadsheet.Protection.Enabled = False
    Spreadsheet.Range("A2").ParseText _
    toto, sDelimiteur$
    
End If
FRM.Charger_Colection Spreadsheet, Mytype

Spreadsheet.AutoFilterMode = False
Spreadsheet.Range("a1").AutoFilter
On Error Resume Next
    Spreadsheet.Range("a1").CurrentRegion.AutoFitColumns
    Err.Clear
    Spreadsheet.Range("a1").CurrentRegion.Cells.EntireColumn.AutoFit
    Err.Clear
    On Error GoTo 0
    For I = 0 To Rs.Fields.Count - 1
   
       
        If Rs.Fields(I).Type = adBoolean Then
        
        DoEvents
            FrmApelan.ProgressBar1.Value = 0
            FrmApelan.ProgressBar1.Max = Spreadsheet.Cells(1, I + 1).CurrentRegion.Rows.Count
            For L = 2 To Spreadsheet.Cells(1, I + 1).CurrentRegion.Rows.Count
                IncremanteBarGrah FrmApelan
                 Spreadsheet.Cells(L, I + 1).Value = Replace(Spreadsheet.Cells(L, I + 1).Value, "'", "")
            Next
            Spreadsheet.Columns(I + 1).NumberFormat = "Yes/No"
        End If
         If InStr(UCase(Rs.Fields(I).Name), UCase("Prix Total")) <> 0 Then
           FrmApelan.ProgressBar1.Value = 0
            FrmApelan.ProgressBar1.Max = Spreadsheet.Cells(1, I + 1).CurrentRegion.Rows.Count
            For L = 2 To Spreadsheet.Cells(1, I + 1).CurrentRegion.Rows.Count
                IncremanteBarGrah FrmApelan
                
                Spreadsheet.Cells(L, I + 1).Formula = Replace(Spreadsheet.Cells(L, I + 1).Value, "'", "")
            Next
         End If
        
    Next
    FrmApelan.ProgressBar1.Value = 0
  If Save_ProgressBar1Max = 0 Then Save_ProgressBar1Max = Save_ProgressBar1Value + 1
 FrmApelan.ProgressBar1.Max = Save_ProgressBar1Max
  FrmApelan.ProgressBar1.Value = Save_ProgressBar1Value
  

End Sub

Function ChercheXls(Val, Myrange2, Optional PasExcel As Boolean, Optional UneFoix As Boolean, Optional Start As Long = 1) As Long
ChercheXls = 1
If Start > 1 Then ChercheXls = Start

Dim RowSave As Long
Dim RageTrouve
ReTante:

On Error Resume Next
If PasExcel = False Then
 Set RageTrouve = Myrange2.Find(What:=Val, After:=Myrange2.Cells(ChercheXls, 1), LookIn:=xlFormulas, _
        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)

Else
Set RageTrouve = Myrange2.Cells.Find(Val, Myrange2.Cells(ChercheXls, 1), ssValues, ssPart)  'Myrange2.Find(Val, Myrange2.Cells(1, 1), Myrange2.xlValues, Myrange2.xlPart)
End If
ChercheXls = 0
ChercheXls = RageTrouve.Row
DoEvents
If Err Then
Err.Clear
    ChercheXls = 0
    On Error GoTo 0
    GoTo Fin
End If
If RowSave > ChercheXls Then
    ChercheXls = RowSave
        On Error GoTo 0
    GoTo Fin
End If
If UneFoix = True Then GoTo Fin
If RowSave = ChercheXls Then

GoTo Fin
End If
If UCase(Trim("" & RageTrouve)) <> UCase(Trim("" & Val)) Then
If Trim("" & RageTrouve) = "" Then
    ChercheXls = 0
    GoTo Fin
End If
RowSave = ChercheXls
ChercheXls = ChercheXls + 1
Set RageTrouve = Nothing
On Error GoTo 0

GoTo ReTante
End If

' For I = 2 To Myrange.Count
'                If UCase(Trim("" & Myrange(I))) = UCase(Trim("" & Val)) Then
'                If Cherche2 = True Then
'                    If Myrange2(I) = 1 Then
'                         ChercheXls = I
'                        Exit For
'                    End If
'                Else
'                        ChercheXls = I
'                        Exit For
'                End If
'                End If
'            Next I

Fin:
On Error GoTo 0
End Function
Function ReplaceHtml(Txt)
ReplaceHtml = Replace(Txt, Chr(10), "<br>")
End Function
Public Sub SendMal(Routine As String, Pj As String)
Dim Rs As Recordset
Dim Sql As String
Dim Destinataire As String
Dim Sujet As String
Dim Body As String
Sql = "SELECT T_Message_Mail.Sujet,T_Message_Mail.Body, T_Users.Email "
Sql = Sql & "FROM T_Users INNER JOIN (T_Message_Mail INNER JOIN T_Destinataire ON T_Message_Mail.Id = "
Sql = Sql & "T_Destinataire.Id_Message) ON T_Users.Id = T_Destinataire.Id_Useur "
Sql = Sql & "WHERE T_Message_Mail.Routine='" & MyReplace(Routine) & " ' "
Sql = Sql & "AND T_Users.Email Is Not Null;"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = False Then
    While Rs.EOF = False
        Destinataire = Destinataire & Rs!Email & ";"
        Sujet = Rs!Sujet
        Body = ReplaceHtml(Rs!Body)
        Rs.MoveNext
    Wend
    Destinataire = Left(Destinataire, Len(Destinataire) - 1)
    Set Rs = Con.OpenRecordSet("SELECT T_Serveur_Smtp.* FROM T_Serveur_Smtp;")
    MailEnvoi Rs!SMTP, Rs!Authentification, Rs!Utilisatuer, Rs!PassWord, Rs!Port, 15, Rs!Messagerie, Destinataire, "", Sujet, Body, Pj

End If



End Sub
'Public reg As ZebClass
Public Sub MailEnvoi(Serveur, Identify As Boolean, User, PassWord, Port, Delay, Expediteur, Dest, DestEnCopy, Objet, Body, Pj)
' sub pour envoyer les mails
Dim msg
Dim Conf
Dim Config
Dim ess
Set msg = CreateObject("CDO.Message") 'pour la configuration du message
Set Conf = CreateObject("CDO.Configuration") '  pour la configuration de l'envoi
Dim strHTML
Set Config = Conf.Fields

' Configuration des parametres d'envoi
'(SMTP - Identification - SSL - Password - Nom Utilisateur - Adresse messagerie)
With Config
If Identify = False Then GoTo Anon
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
    .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = User
    .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = PassWord
Anon:
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = Port
    .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = Serveur
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = Delay
    .Update
End With
DoEvents

'Configuration du message
'If E_mail.Sign.Value = Checked Then Convert ServeurFrm.SignTXT, ServeurFrm.Text1

With msg
    Set .Configuration = Conf
    .To = Dest
  .cc = DestEnCopy
    .FROM = Expediteur
    .Subject = Objet
'            If E_mail.Sign.Value = 1 Then _
    .htmlbody = E_mail.ZThtml.Text & "<p align=""left""><font face=""MS Sans Serif"" size=""1"" color=""#000000""><b>" & "---------------------------------------" & "<P></P>" & ServeurFrm.Text1.Text _
            Else _
.sender"toto"

    .htmlbody = Body '"<p align=""center""><font face=""Verdana"" size=""1"" color=""#9224FF""><b><br><font face=""Comic Sans MS"" size=""5"" color=""#FF0000""></b><i>" & body & "</i></font> " 'E_mail.ZThtml.Text
            If Pj <> "" Then _
    .AddAttachment Pj
    .Send 'envoi du message

End With
DoEvents
' reinitialisation des variables
Set msg = Nothing
Set Conf = Nothing
Set Config = Nothing

DoEvents
End Sub

Function DecodeCode_APP(Code_APP As String) As String
Dim SplitCode_APP
Dim NbUbound As Long
SplitCode_APP = Split(Code_APP, ".")
NbUbound = UBound(SplitCode_APP)
Select Case NbUbound
Case -1
    DecodeCode_APP = vbNullChar
Case 0
    DecodeCode_APP = SplitCode_APP(0)
Case 1
    DecodeCode_APP = SplitCode_APP(0)
Case 2
    DecodeCode_APP = SplitCode_APP(1)
Case Else
     DecodeCode_APP = SplitCode_APP(1)
    End Select
End Function
 Function MyFormat(Mytype As String, MyText As Object, MyLib As String) As Boolean
 MyFormat = True
 If MyText = "" Then Exit Function
 
  Select Case UCase(Mytype)
                    Case "DATE"
                        If Not IsDate(MyText) Then
                            MsgBox "Vous devez saisir une date pour : " & MyLib, vbExclamation
                            MyText = ""
                            MyFormat = False
                            
                            Exit Function
                        Else
                            MyText = Format(MyText, "dd/mm/yyyy")
                        End If
                    Case "ENT"
                        If Not IsNumeric(MyText) Then
                            MsgBox "Vous devez saisir un nombre entier pour : " & MyLib, vbExclamation
                            MyText = ""
                            MyFormat = False
                            Exit Function
                        Else
                            If (InStr(1, (MyText), ",") <> 0) Or (InStr(1, (MyText), ".") <> 0) Then
                                MsgBox "Vous devez saisir un nombre entier pour : " & MyLib, vbExclamation
                                MyText = ""
                                MyFormat = False
                                Exit Function
                            End If
                        End If
                    Case "DBL"
                        If Not IsNumeric(MyText) Then
                            MyText = Replace(MyText, ".", ",")
                        End If
                        If Not IsNumeric(MyText) Then
                            MsgBox "Vous devez saisir un nombre à virgule pour : " & MyLib, vbExclamation
                            MyText = ""
                            MyFormat = False
                            Exit Function
                        End If
            End Select
 End Function
Sub Main()
If Trim("" & Command) <> "[N ais pas peur fils papa est la]" Then
    MsgBox "Ce module ne peut être exécuté qu'à partir d'une licence Autocâble.", vbCritical
    End
End If
LoadDb
Modifier.Show

End Sub
Public Sub ReplaceTousXls(MySeet As Worksheet, Recherche As String, Ramplace As String)
'
' Macro1 Macro
' Macro enregistrée le 14/11/2006 par robert.durupt
'

'
    MySeet.Cells.Replace What:=Recherche, Replacement:=Ramplace, LookAt:= _
        xlWhole, SearchOrder:=xlByColumns, MatchCase:=False, SearchFormat:=False _
        , ReplaceFormat:=False
End Sub

Function CherCheInFihier(Cherher As String) As String
Dim FileNumber As Long
Dim MyString As String
Dim Spliligne
FileNumber = FreeFile

  
Open App.Path & "\Autocable.ini" For Input As #FileNumber
Do While Not EOF(FileNumber)    ' Effectue la boucle jusqu'à la fin du fichier.
    Input #FileNumber, MyString ' Lit les données dans deux variables.
    If InStr(1, MyString, Cherher) <> 0 Then
    Spliligne = Split(MyString & "====", "=")
       CherCheInFihier = Trim(Spliligne(1))
'       CherCheInFihier = Right(MyString, Len(MyString) - (Len(Cherher) + InStr(1, MyString, Cherher)))
       Exit Do
    End If
Loop
Close #FileNumber    ' Ferme le fichier
CherCheInFihier = Replace(CherCheInFihier, "§remplaceDate§", Format(Date, "yyyy"))
CherCheInFihier = Trim(CherCheInFihier)
End Function
Sub LoadDb()
BdDateTable = CherCheInFihier("BdDateTable")
DbNumPlan = CherCheInFihier("Bdnumero")
If UCase(CherCheInFihier("IsCilent")) = "TRUE" Then IsCilent = True

If UCase(CherCheInFihier("IsServeur")) = "TRUE" Then IsServeur = True
Db = CherCheInFihier("BdAutocable")
AutocableDRIVE = CherCheInFihier("AutocableDRIVE")
DonneesEntreprise = CherCheInFihier("DonneesEntreprise")
DonneesProduction = CherCheInFihier("DonneesProduction")
Con.BASE = Db
Con.TYPEBASE = 5

Con.OpenConnetion
If IsServeur = IsCilent Then IsServeur = False: IsCilent = False
End Sub
Function funPath()
    Dim MyPath As New Collection
    Dim Rs As Recordset
        Set Rs = Con.OpenRecordSet("SELECT T_Path.* FROM T_Path;")
    While Rs.EOF = False
        MyPath.Add Rs.Fields("PathVar").Value, Rs.Fields("NameVar").Value
        Rs.MoveNext
    Wend
    Set Rs = Con.CloseRecordSet(Rs)
    Set funPath = MyPath
End Function
Public Function SershXls(Feuille As Worksheet, Valeur As String) As Long
On Error Resume Next
SershXls = 0
  SershXls = Feuille.Cells.Find(What:=Valeur, After:=Feuille.Cells(1, 1), LookIn:=xlFormulas, _
        LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Row

End Function
Function ChoixCouleur(Mode As Long, Optional BoolExcel As Boolean) As Long
   
  If BoolExcel = False Then
   Select Case Mode
   Case 0
        ChoixCouleur = 12632256
    Case 1
        ChoixCouleur = 16777164
    Case 2
    ChoixCouleur = 10079487
    Case 3
        ChoixCouleur = 13434828
    Case 4
        ChoixCouleur = &HFFC0FF
   End Select

Else
    Select Case Mode
    Case 0
        ChoixCouleur = 15
    Case 1
        ChoixCouleur = 34
    Case 2
    ChoixCouleur = 40
    Case 3
        ChoixCouleur = 35
    Case 4
        ChoixCouleur = 38
   End Select
End If
End Function
Public Function VersionPices(Pieces As String) As Long
Dim Rs As Recordset
Dim Sql As String
Sql = "SELECT  VersionPices.Version FROM VersionPices "
Sql = Sql & "WHERE VersionPices.Pi='" & MyReplace(Pieces) & "';"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = True Then
Sql = "INSERT INTO VersionPices ( Pi ) VALUES('" & MyReplace(Pieces) & "');"
Con.Execute Sql
End If
Sql = "UPDATE VersionPices SET VersionPices.Version = [Version] + 1 "
Sql = Sql & "WHERE VersionPices.Pi='" & MyReplace(Pieces) & "';"
Con.Execute Sql
Rs.Requery
VersionPices = Rs!Version
End Function
Function MyReplace(strVal As String) As String
strVal = Trim(strVal)
MyReplace = strVal
MyReplace = Replace(MyReplace, "'", "''")
MyReplace = Replace(MyReplace, Chr(34), Chr(34) & Chr(34))
MyReplace = Trim("" & MyReplace)
End Function
Public Sub InsertRow(MyRange As Range, L As Long)
MyRange.Rows(L & ":" & L).Insert Shift:=xlDown
End Sub
Function PathArchive(PathRacicine As String, Client As String, CleAc As String, Piece As String, Mytype As String, Fichier, IdPieces As Long, Indice_Pieces As String, Indice_Plan As String, Version As Long, Optional NoRegistre As Boolean) As String
Dim Fso As New FileSystemObject
Dim Sql As String
Dim MyPath
Dim aa
Dim NomenclatureOk As Boolean
Dim IndexP As Long
Dim Rs As Recordset
Indice_Pieces = Trim("" & Indice_Pieces)
Indice_Plan = Trim("" & Indice_Plan)
Piece = Replace(Piece, "/", "_", 1)
Piece = Replace(Piece, ":", "", 1)
Piece = Replace(Piece, ".", "", 1)
Piece = Piece & "_" & Indice_Pieces
If UCase(Mytype) = UCase("SyntG") Or UCase(Mytype) = UCase("pdf") Or UCase(Mytype) = UCase("Synt") Or Mytype = "LIEC" Or Mytype = "DAC" Or Mytype = "DNC" Or Mytype = "FAB" Then
Else
Fichier = Fichier & "_" & Indice_Plan
End If
Fichier = Replace(Fichier, "/", "_", 1)
Fichier = Replace(Fichier, ":", "", 1)
Fichier = Replace(Fichier, ".", "", 1)



PathArchive = TableauPath.Item(Mytype)
PathArchive = Replace(UCase(PathArchive), UCase("[VariableClient]"), Client)
PathArchive = Replace(UCase(PathArchive), UCase("[VariableAff]"), CleAc)
    PathArchive = Replace(UCase(PathArchive), UCase("[VaribleDoc]"), Fichier)

    


If Version > 1 Then
    
    PathArchive = Replace(UCase(PathArchive), UCase("[VARIABLEPI]"), Piece & "_MOD")
Else
    PathArchive = Replace(UCase(PathArchive), UCase("[VARIABLEPI]"), Piece)
End If

PathRacicine = DefinirChemienComplet(TableauPath.Item("PathServer"), PathRacicine)
MyPath = Split(PathArchive, "\")
aa = ""
For IndexP = 0 To UBound(MyPath) - 1
aa = aa & MyPath(IndexP) & "\"
Debug.Print Left(PathRacicine & "\" & aa, Len(PathRacicine & "\" & aa) - 1)
If Fso.FolderExists(Left(PathRacicine & "\" & aa, Len(PathRacicine & "\" & aa) - 1)) = False Then
    Fso.CreateFolder Left(PathRacicine & "\" & aa, Len(PathRacicine & "\" & aa) - 1)
End If
Next


If NoRegistre = False Then
    If NomenclatureOk = True Then
        Sql = "UPDATE T_indiceProjet SET T_indiceProjet." & UCase(Mytype) & "AutoCadSave = '" & MyReplace(PathArchive) & "', T_indiceProjet.NbErr = " & NbError & ",T_indiceProjet." & UCase(Mytype) & "Ok=true "
        Sql = Sql & "WHERE T_indiceProjet.Id=" & IdPieces & ";"
     Else
        Sql = "UPDATE T_indiceProjet SET T_indiceProjet." & UCase(Mytype) & "AutoCadSave = '" & MyReplace(PathArchive) & "' "
        Sql = Sql & "WHERE T_indiceProjet.Id=" & IdPieces & ";"
     End If
    Con.Execute Sql
End If
    


PathArchive = PathRacicine & "\" & PathArchive
Debug.Print PathArchive
End Function
Function DefinirChemienComplet(Serveur As String, Path As String) As String
If Right(Trim("" & Serveur), 1) <> "\" Then Serveur = Serveur & "\"
If Trim("" & Path) = "" Then
    DefinirChemienComplet = Serveur
Else
    If Left(Path, 1) = "\" And Left(Path, 2) <> "\\" Then Path = Right(Path, Len(Path) - 1)
DefinirChemienComplet = Path
End If
If Mid(DefinirChemienComplet, 2, 1) = ":" Then Exit Function
If Left(Path, 1) <> "\" Then
    If Right(Serveur, 1) <> "\" Then
        DefinirChemienComplet = Serveur & "\" & DefinirChemienComplet
    Else
         DefinirChemienComplet = Serveur & DefinirChemienComplet
    End If
End If
If Right(Trim(DefinirChemienComplet), 2) = "\\" Then DefinirChemienComplet = Mid(DefinirChemienComplet, 1, Len(DefinirChemienComplet) - 1)
If Left(DefinirChemienComplet, 1) = "\" And Left(DefinirChemienComplet, 2) <> "\\" Then DefinirChemienComplet = "\" & DefinirChemienComplet
Debug.Print DefinirChemienComplet
End Function
