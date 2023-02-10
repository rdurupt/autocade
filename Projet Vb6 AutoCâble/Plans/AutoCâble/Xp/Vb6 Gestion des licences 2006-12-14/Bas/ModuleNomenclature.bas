Attribute VB_Name = "ModuleNomenclature"
Function Nomenclature2(Id_IndiceProjet As Long, Optional PathPl As String, Optional Save As Boolean = True) As Boolean
Dim ConConnecteur As New Ado
 Dim TbAlve As New Collection
Dim Sql As String
Dim Rs As Recordset
Dim Rs2 As Recordset
Dim NumFieldsConnecteur As Long
Dim NumFieldsFournisseur As Long
Dim NumFieldsBouchon As Long
Dim NumFieldsCapot As Long
Dim NumFieldsVerou As Long
Dim NumFieldsJoint As Long
Dim NumFieldsAlveole As Long
Dim DbOk As Boolean
Dim Id As Long
Dim MyEtiquette As New ClsEtiqette
Dim MyRange As Range
Dim PathModelWord As String
Dim PIE(1) As String
Dim Ensemble(1) As String
Dim I As Long
Dim RsComposant As Recordset
Dim RsFils As Recordset
Dim AlveolOk As Boolean
Dim iTotal As Long
Dim MySeet As Worksheet
Set MySeet = IsertSheet(MyWorkbook, "Appro Connectique", True)
'MySeet.Application'.Visible = True
ConConnecteur.TYPEBASE = ADO_TYPEBASE
ConConnecteur.SERVER = ADO_SERVER
ConConnecteur.User = ADO_User
ConConnecteur.PassWord = ADO_PassWord
ConConnecteur.BASE = DbCatalogue
ConConnecteur.OpenConnetion

DbOk = ConConnecteur.OpenConnetion '"U:\Librairies\Plans\AutoCâble\Xp\Access\Catalogue Renault.mdb"
DeleteRow MySeet, True

Sql = "SELECT [PI] & '_' & [PI_Indice] AS PIE, T_indiceProjet.Ensemble "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Id_IndiceProjet & ";"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = False Then
    PIE(0) = "PI": PIE(1) = "" & Rs!PIE
    Ensemble(0) = "Ensemble": Ensemble(1) = "" & Rs!Ensemble
End If

Sql = "SELECT Rq_Compte_Connecteur_IdPices.CONNECTEUR, Rq_Compte_Connecteur_IdPices.OPTION, Rq_Compte_Connecteur_IdPices.[Qté],0 as [Prix U],'=(RC[-1]*RC[-2]) ' as [Prix Total] "
Sql = Sql & "FROM Rq_Compte_Connecteur_IdPices "
Sql = Sql & "WHERE Rq_Compte_Connecteur_IdPices.CONNECTEUR<>'NEANT' "
Sql = Sql & "AND Rq_Compte_Connecteur_IdPices.Id_IndiceProjet=" & Id_IndiceProjet & " " '94 "
Sql = Sql & "ORDER BY Rq_Compte_Connecteur_IdPices.CONNECTEUR;"
Set Rs = Con.OpenRecordSet(Sql)
NumFieldsConnecteur = Rs.Fields.Count
    If Rs.EOF = True Then Exit Function
    
Set MyWord = CreateObject("Word.Application")



PathModelWord = TableauPath.Item("PathModelWord")
PathModelWord = DefinirChemienComplet(TableauPath.Item("PathServer"), PathModelWord)
'         If Left(PathModelWord, 2) <> "\\" And Left(PathModelWord, 1) = "\" Then PathModelWord = TableauPath.Item("PathServer") & PathModelWord
'          If Right(PathModelWord, 2) = "\\" Then PathModelWord = Mid(PathModelWord, 1, Len(PathModelWord) - 1)
Set MyWordDoc = WordNewDoc(PathModelWord)
          
          
PathModelWord = TableauPath.Item("PathModelWordMarc")
PathModelWord = DefinirChemienComplet(TableauPath.Item("PathServer"), PathModelWord)
'         If Left(PathModelWord, 2) <> "\\" And Left(PathModelWord, 1) = "\" Then PathModelWord = TableauPath.Item("PathServer") & PathModelWord
'          If Right(PathModelWord, 2) = "\\" Then PathModelWord = Mid(PathModelWord, 1, Len(PathModelWord) - 1)
Set MyWordDoc2 = WordNewDoc(PathModelWord)
          
'MyExcel'.Visible = True
    
While Rs.EOF = False
NbLigne = NbLigne + 1
Rs.MoveNext
Wend
Rs.Requery
'MyWorkbook.Application'.Visible = True
For I = 0 To Rs.Fields.Count - 1
'MySeet.Application'.Visible = True
    MySeet.Cells(5, I + 1) = Rs(I).Name
Next
Row = 5
 FormBarGrah.ProgressBar1.Value = 0
 If NbLigne = 0 Then NbLigne = 1
 FormBarGrah.ProgressBar1.Max = NbLigne
 FormBarGrah.ProgressBar1Caption.Caption = " Exporter liste des Appros :"
DoEvents
While Rs.EOF = False
Row = Row + 1
  IncremanteBarGrah FormBarGrah
   Connecteur = "" & Rs(0).Value
   If InStr(1, "" & Connecteur, "§") <> 0 Then
    Connecteur = Split(Connecteur, "§")
    Connecteur = Connecteur(0)
   End If
    MySeet.Cells(Row, 1) = Connecteur
   MySeet.Cells(Row, 2) = "" & Rs(1)
   MySeet.Cells(Row, 3) = "" & Rs(2)
      MySeet.Cells(Row, 4) = "" & Rs(3)
      MySeet.Cells(Row, 5).FormulaR1C1 = "" & Rs(4)
      Sql = "SELECT Connecteurs.CODE_APP, Connecteurs.DESIGNATION "
    Sql = Sql & "FROM Connecteurs "
    Sql = Sql & "WHERE Connecteurs.Id_IndiceProjet=" & Id_IndiceProjet & " "
    Sql = Sql & "AND Connecteurs.CONNECTEUR='" & Connecteur & "' "
    If Trim(MyReplace("" & Rs!Option)) = "" Then
        Sql = Sql & "AND Connecteurs.[O/N]=False AND Connecteurs.OPTION is null ;"
    Else
        Sql = Sql & "AND Connecteurs.[O/N]=False AND Connecteurs.OPTION='" & MyReplace("" & Rs!Option) & "' ;"
    End If
     Set Rs2 = Con.OpenRecordSet(Sql)
        NumFieldsConnecteurApp = NumFieldsConnecteur + Rs2.Fields.Count - 1
     MyEtiquette.PrpareEtiqet Rs2.GetString, ""
     MyEtiquette.RenseigneChamp "Connecteur", "" & Connecteur
     MyEtiquette.RenseigneChamp "" & PIE(0), "" & PIE(1)
     MyEtiquette.RenseigneChamp "" & Ensemble(0), "" & Ensemble(1)
    For I = 1 To Rs2.Fields.Count
    
        MySeet.Cells(5, I + NumFieldsConnecteur) = Rs2(I - 1).Name
    
    Next
    Rs2.Requery
    If Rs2.EOF = True Then
  
        FunError 8, "" & Connecteur, "" & vbCrLf & Err.Description
    End If
    
    While Rs2.EOF = False
    
            For I = 1 To Rs2.Fields.Count
        MySeet.Cells(Row, I + NumFieldsConnecteur) = MySeet.Cells(Row, I + NumFieldsConnecteur) & Chr(10) & Replace("" & Rs2(I - 1), Chr(13), "")
         
        Next
             Rs2.MoveNext
        Wend
   If DbOk = True Then
    Sql = "SELECT Rq_Fournisseur.* "
    Sql = Sql & "FROM Rq_Fournisseur "
    Sql = Sql & "WHERE Rq_Fournisseur.[Ref Connecteur]= '" & Connecteur & "';"
    Set Rs2 = ConConnecteur.OpenRecordSet(Sql)
    NumFieldsFournisseur = NumFieldsConnecteurApp + Rs2.Fields.Count - 1
    For I = 2 To Rs2.Fields.Count - 1
        MySeet.Cells(5, I + NumFieldsConnecteurApp) = Rs2(I).Name
    
    Next
    If Rs2.EOF = True Then
  
        FunError 8, "" & Connecteur, "" & vbCrLf & Err.Description
    End If
    While Rs2.EOF = False
    
            For I = 2 To Rs2.Fields.Count - 1
        MySeet.Cells(Row, I + NumFieldsConnecteurApp) = MySeet.Cells(Row, I + NumFieldsConnecteurApp) & Chr(10) & Replace("" & Rs2(I), Chr(13), "")
         
        Next
             Rs2.MoveNext
        Wend
        
        
        Sql = "SELECT Rq_Bouchon.* "
    Sql = Sql & "FROM Rq_Bouchon "
    Sql = Sql & "WHERE Rq_Bouchon.Référence= '" & Connecteur & "';"
    Set Rs2 = ConConnecteur.OpenRecordSet(Sql)
   

    NumFieldsBouchon = NumFieldsFournisseur + Rs2.Fields.Count - 1
    For I = 1 To Rs2.Fields.Count - 1
          MySeet.Cells(5, I + NumFieldsFournisseur) = Replace(Replace(Replace(Rs2(I).Name, "Qté", "Bouchon Qté"), "Prix U", "Bouchon Prix U"), "Prix Total", "Bouchon Prix Total")
         
         If Rs2(I).Name = "Qté" Or Rs2(I).Name = "Prix U" Then
          
            MySeet.Cells(Row, I + NumFieldsFournisseur) = 0
         End If
         If Rs2(I).Name = "Prix Total" Then
            MySeet.Cells(Row, I + NumFieldsFournisseur).FormulaR1C1 = "=(RC[-1]*RC[-2])"
         End If
    Next
     While Rs2.EOF = False
   
            For I = 1 To Rs2.Fields.Count - 1
                  
                   
                             If Rs2(I).Name = "Qté" Or Rs2(I).Name = "Prix U" Or Rs2(I).Name = "Prix Total" Then
                                
                              Else
                                MySeet.Cells(Row, I + NumFieldsFournisseur) = MySeet.Cells(Row, I + NumFieldsFournisseur) & Chr(10) & Replace("" & Rs2(I), Chr(13), "")
                  End If
                  Next
             Rs2.MoveNext
        Wend
        
           
        Sql = "SELECT Rq_Capot.* "
    Sql = Sql & "FROM Rq_Capot "
    Sql = Sql & "WHERE Rq_Capot.Référence= '" & Connecteur & "';"
    Set Rs2 = ConConnecteur.OpenRecordSet(Sql)
  
    NumFieldsCapot = NumFieldsBouchon + Rs2.Fields.Count - 1
    For I = 1 To Rs2.Fields.Count - 1
          MySeet.Cells(5, I + NumFieldsBouchon) = Rs2(I).Name
         
    Next
     While Rs2.EOF = False
   
            For I = 0 To Rs2.Fields.Count - 1
          MySeet.Cells(Row, I + NumFieldsBouchon) = MySeet.Cells(Row, I + NumFieldsBouchon) & Chr(10) & Replace("" & Rs2(I), Chr(13), "")
         
        Next
             Rs2.MoveNext
        Wend
        
      Sql = "SELECT Rq_Verou.* "
    Sql = Sql & "FROM Rq_Verou "
    Sql = Sql & "WHERE Rq_Verou.Référence= '" & Connecteur & "';"
    Set Rs2 = ConConnecteur.OpenRecordSet(Sql)
    
   
   NumFieldsVerou = NumFieldsCapot + Rs2.Fields.Count - 1
    For I = 1 To Rs2.Fields.Count - 1
          MySeet.Cells(5, I + NumFieldsCapot) = Rs2(I).Name
         
    Next
     While Rs2.EOF = False
   
            For I = 0 To Rs2.Fields.Count - 1
          MySeet.Cells(Row, I + NumFieldsCapot) = MySeet.Cells(Row, I + NumFieldsCapot) & Chr(10) & Replace("" & Rs2(I), Chr(13), "")
         
        Next
             Rs2.MoveNext
        Wend
         Sql = "SELECT Rq_Joint.* "
    Sql = Sql & "FROM Rq_Joint "
    Sql = Sql & "WHERE Rq_Joint.Référence= '" & Connecteur & "';"
    Set Rs2 = ConConnecteur.OpenRecordSet(Sql)
  
  NumFieldsJoint = NumFieldsVerou + Rs2.Fields.Count - 1
    For I = 1 To Rs2.Fields.Count - 1
          MySeet.Cells(5, I + NumFieldsVerou) = Replace(Replace(Replace(Rs2(I).Name, "Qté", "Joint Qté"), "Prix U", "Joint Prix U"), "Prix Total", "Joint Prix Total")
           If Rs2(I).Name = "Qté" Or Rs2(I).Name = "Prix U" Then
            MySeet.Cells(Row, I + NumFieldsVerou) = 0
           End If
           If Rs2(I).Name = "Prix Total" Then
           MySeet.Cells(Row, I + NumFieldsVerou).FormulaR1C1 = "=(RC[-1]*RC[-2])"
           End If
         
    Next
     While Rs2.EOF = False
   
            For I = 1 To Rs2.Fields.Count - 1
          
          
                 If Rs2(I).Name = "Qté" Or Rs2(I).Name = "Prix U" Or Rs2(I).Name = "Prix Total" Then
                   
                Else
                MySeet.Cells(Row, I + NumFieldsVerou) = MySeet.Cells(Row, I + NumFieldsVerou) & Chr(10) & Replace("" & Rs2(I), Chr(13), "")
                End If
            
          MyEtiquette.RenseigneChamp "" & Rs2(I).Name, "" & Rs2(I).Value
        Next
             Rs2.MoveNext
'        End If
Wend
DoEvents
AlveolOk = False
'If "7703297847" = Connecteur Then
'MsgBox ""
'End If

Sql = "DELETE TempFillesSection.* FROM TempFillesSection "
Sql = Sql & "WHERE TempFillesSection.Id_IndiceProjet=" & Id_IndiceProjet & ";"
ConConnecteur.Execute Sql
Con.Execute Sql

 Sql = "SELECT TempFillesSection.* FROM TempFillesSection "
    Sql = Sql & "WHERE TempFillesSection.Id_IndiceProjet=" & Id_IndiceProjet & ";"

  Set RsFils = Con.OpenRecordSet(Sql)
  
  
  

    Sql = "SELECT Ligne_Tableau_fils.Id_IndiceProjet, Ligne_Tableau_fils.SECT, Ligne_Tableau_fils.VOI AS Vois, Ligne_Tableau_fils.APP AS Cod_App,  "
    Sql = Sql & "Connecteurs.CONNECTEUR "
    Sql = Sql & "FROM Ligne_Tableau_fils INNER JOIN Connecteurs ON (Ligne_Tableau_fils.APP = Connecteurs.CODE_APP)  "
    Sql = Sql & "AND (Ligne_Tableau_fils.Id_IndiceProjet = Connecteurs.Id_IndiceProjet) "
    Sql = Sql & "GROUP BY Ligne_Tableau_fils.Id_IndiceProjet, Ligne_Tableau_fils.SECT, Ligne_Tableau_fils.VOI,  "
    Sql = Sql & "Ligne_Tableau_fils.APP, Connecteurs.CONNECTEUR "
    Sql = Sql & "Having Ligne_Tableau_fils.Id_IndiceProjet = " & Id_IndiceProjet & " And Ligne_Tableau_fils.VOI Is Not Null  "
    Sql = Sql & "And Ligne_Tableau_fils.App Is Not Null And Connecteurs.Connecteur =  '" & Connecteur & "' "
    Sql = Sql & "ORDER BY Ligne_Tableau_fils.APP;"
DoEvents
Set Rs2 = Con.OpenRecordSet(Sql)
DoEvents
While Rs2.EOF = False
RsFils.AddNew
RsFils!Id_IndiceProjet = Rs2!Id_IndiceProjet
RsFils!SECT = ConvertTxtAsDouble("" & Rs2!SECT)
RsFils!VOI = "" & Rs2!VOIS
RsFils!App = "" & Rs2!Cod_App
Connecteur = "" & Rs2!Connecteur
RsFils.Update
DoEvents
    Rs2.MoveNext
Wend
  Set Rs2 = Con.CloseRecordSet(Rs2)

        Sql = "SELECT Ligne_Tableau_fils.Id_IndiceProjet, Ligne_Tableau_fils.SECT, Ligne_Tableau_fils.VOI2 AS Vois, "
    Sql = Sql & "Ligne_Tableau_fils.APP2 AS Cod_App, Connecteurs.CONNECTEUR "
    Sql = Sql & "FROM Ligne_Tableau_fils INNER JOIN Connecteurs ON (Ligne_Tableau_fils.APP2 = Connecteurs.CODE_APP) "
    Sql = Sql & "AND (Ligne_Tableau_fils.Id_IndiceProjet = Connecteurs.Id_IndiceProjet) "
    Sql = Sql & "GROUP BY Ligne_Tableau_fils.Id_IndiceProjet, Ligne_Tableau_fils.SECT, Ligne_Tableau_fils.VOI2,  "
    Sql = Sql & "Ligne_Tableau_fils.APP2, Connecteurs.CONNECTEUR "
    Sql = Sql & "Having Ligne_Tableau_fils.Id_IndiceProjet = " & Id_IndiceProjet & " And Ligne_Tableau_fils.VOI2 Is Not Null "
    Sql = Sql & "And Ligne_Tableau_fils.APP2 Is Not Null And Connecteurs.Connecteur =  '" & Connecteur & "' "
    Sql = Sql & "ORDER BY Ligne_Tableau_fils.APP2;"

DoEvents
Set Rs2 = Con.OpenRecordSet(Sql)
DoEvents
While Rs2.EOF = False
RsFils.AddNew
RsFils!Id_IndiceProjet = Rs2!Id_IndiceProjet
RsFils!SECT = ConvertTxtAsDouble("" & Rs2!SECT)
RsFils!VOI = "" & Rs2!VOIS
RsFils!App = "" & Rs2!Cod_App
Connecteur = "" & Rs2!Connecteur
RsFils.Update
DoEvents
    Rs2.MoveNext
Wend
 DoEvents
    Sql = "INSERT INTO TempFillesSection ( Id_IndiceProjet, VOI, SECT, APP ) IN '" & ConConnecteur.RetournDbName("MDB") & "' "
    Sql = Sql & "SELECT TempFillesSection.Id_IndiceProjet, TempFillesSection.VOI, Sum(TempFillesSection.SECT) AS SommeDeSECT, TempFillesSection.APP "
    Sql = Sql & "FROM TempFillesSection "
    Sql = Sql & "GROUP BY TempFillesSection.Id_IndiceProjet, TempFillesSection.VOI, TempFillesSection.APP "
    Sql = Sql & "HAVING TempFillesSection.Id_IndiceProjet=" & Id_IndiceProjet & ";"
    DoEvents
Con.Execute Sql
 MySeconde 2
'
'      sql = "SELECT Rq_Alveole.* "
'    sql = sql & "FROM Rq_Alveole "
'    sql = sql & "WHERE Rq_Alveole.Référence= '" & Connecteur & "';"
'    Set Rs2 = ConConnecteur.OpenRecordSet(sql)
'
DoEvents
    
    
    Sql = "SELECT Rq_Alveole.Référence, Rq_Alveole.[Nb Alvé], Rq_Alveole.Voie, Rq_Alveole.Famille, Rq_Alveole.[Famille Lib],  "
    Sql = Sql & "Rq_Alveole.[Alvé Réf], Rq_Alveole.Qté, Rq_Alveole.[Prix U], Rq_Alveole.[Prix Total], Rq_Alveole.[Alvé Réf Fourr],  "
    Sql = Sql & "Rq_Alveole.[Alvéole Mini en mm2], Rq_Alveole.[Alvéole Maxi en mm2] "
    Sql = Sql & "FROM Rq_Alveole "
    Sql = Sql & "WHERE Rq_Alveole.Id_IndiceProjet=" & Id_IndiceProjet & " AND Rq_Alveole.Référence='" & Connecteur & "' "
    Sql = Sql & "GROUP BY Rq_Alveole.Référence, Rq_Alveole.[Nb Alvé], Rq_Alveole.Voie, Rq_Alveole.Famille, Rq_Alveole.[Famille Lib],  "
    Sql = Sql & "Rq_Alveole.[Alvé Réf], Rq_Alveole.Qté, Rq_Alveole.[Prix U], Rq_Alveole.[Prix Total], Rq_Alveole.[Alvé Réf Fourr],  "
    Sql = Sql & "Rq_Alveole.[Alvéole Mini en mm2], Rq_Alveole.[Alvéole Maxi en mm2], Rq_Alveole.Id_IndiceProjet, Rq_Alveole.Référence; "
    DoEvents
 Set Rs2 = ConConnecteur.OpenRecordSet(Sql)
 DoEvents
'    MyWorkbook.Application'.Visible = True
 
   Set TableauAlve = Nothing
   
 DoEvents
   If Rs2.EOF = False Then
        AlveolOk = True
        TableauAlve = Rs2.GetRows
        DoEvents
        ReDim tableauAlve2(UBound(TableauAlve), 1)
        DoEvents
        I = 0
        For I = 0 To UBound(TableauAlve, 2)
            On Error Resume Next
            a = ""
            a = TbAlve(TableauAlve(3, I))
            If Err <> 0 Then
                Err.Clear
                TbAlve.Add I, TableauAlve(3, I)
            End If
            tableauAlve2(TbAlve(TableauAlve(3, I)), 0) = "" & TableauAlve(3, I) & ": "
             tableauAlve2(TbAlve(TableauAlve(3, I)), 1) = tableauAlve2(TbAlve(TableauAlve(3, I)), 1) & "" & TableauAlve(5, I) & "(_____), "
             DoEvents
        Next
        Txt = ""
        For I = LBound(tableauAlve2) To UBound(tableauAlve2)
            If Trim("" & tableauAlve2(I, 1)) <> "" Then
           Txt = Txt & tableauAlve2(I, 0) & tableauAlve2(I, 1) & ";"
           
            Debug.Print Txt
            
            End If
            DoEvents
        Next
    Else
        Txt = ""
        ReDim tableauAlve2(0, 1)
     End If

Txt = Replace(Txt, ",;", "; ")
Txt = Replace(Txt, ", ;", "; ")
Debug.Print Txt
'txt = Replace(txt, ":", "")
MyEtiquette.RenseigneChamp "Famille", "" & Txt
' MySeet.Application'.Visible = True
'  ReDim tableauAlve(1, 1) Famille
  Dim T_Alve() As String
' ReDim T_Alve(Bound(tableauAlve), 1)
  Rs2.Requery
  NumFieldsAlveole = NumFieldsJoint + Rs2.Fields.Count - 1
    For I = 1 To Rs2.Fields.Count - 1
        If iTotal < I + NumFieldsJoint Then iTotal = I + NumFieldsJoint
          MySeet.Cells(5, I + NumFieldsJoint) = Replace(Replace(Replace(Rs2(I).Name, "Qté", "Alvé Qté"), "Prix U", "Alvé Prix U"), "Prix Total", "Alvé Prix Total")
           If Rs2(I).Name = "Qté" Or Rs2(I).Name = "Prix U" Then
           MySeet.Cells(Row, I + NumFieldsJoint) = 0
         End If
         If Rs2(I).Name = "Prix Total" Then
            MySeet.Cells(Row, I + NumFieldsJoint) = "=(RC[-1]*RC[-2])"
         End If
         
    Next
     While Rs2.EOF = False
   
            For I = 1 To Rs2.Fields.Count - 1
            
                If Rs2(I).Name = "Qté" Or Rs2(I).Name = "Prix U" Or Rs2(I).Name = "Prix Total" Then
                Else
                    If UCase(Rs2(I).Name) = UCase("Voie") Then
                        MySeet.Cells(Row, I + NumFieldsJoint) = MySeet.Cells(Row, I + NumFieldsJoint) + Val(Replace("" & Rs2(I - 1), Chr(13), ""))
                    Else
                    MySeet.Cells(Row, I + NumFieldsJoint) = MySeet.Cells(Row, I + NumFieldsJoint) & Chr(10) & Replace("" & Rs2(I), Chr(13), "")
                    End If
                End If
        Next
           
             Rs2.MoveNext
        Wend
End If
      DoEvents
'      MyWord.Application'.Visible = True
On Error GoTo Fin
      For I = MyEtiquette.TableMin To MyEtiquette.TableMax
      a = MyEtiquette.RetournEtiquette(I)
        CreerEtiquette MyEtiquette.RetournEtiquette(I)
      Next
Fin:
'      MyWord'.Visible = True
'      CreerEtiquette
    Rs.MoveNext
Wend

Set Rs2 = ConConnecteur.CloseRecordSet(Rs2)
ConConnecteur.CloseConnection
Set Rs = Con.CloseRecordSet(Rs)

Sql = "SELECT  T_indiceProjet.Ensemble, T_indiceProjet.Equipement,[Li] & '_' & [LI_Indice] AS Liste, "
Sql = Sql & "[PI] & '_' & [PI_Indice] AS Piece, T_indiceProjet.Client, T_indiceProjet.CleAc "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Id_IndiceProjet & ";"
Dim RsEntetePage As Recordset
Set RsEntetePage = Con.OpenRecordSet(Sql)
Set MyRange = MySeet.Range("A5").CurrentRegion
'MyWorkbook.Application'.Visible = True
MySeet.Range("A2") = "TOTAL"
MySeet.Range("D2") = "SOUS TOTAL"
MySeet.Range("N2") = "SOUS TOTAL"
MySeet.Range("W2") = "SOUS TOTAL"
MySeet.Range("AH2") = "SOUS TOTAL"

FormatExcelPlage MySeet.Range("A2"), 15, False, True, xlCenter, xlCenter
FormatExcelPlage MySeet.Range("D2"), 15, False, True, xlCenter, xlCenter
FormatExcelPlage MySeet.Range("N2"), 15, False, True, xlCenter, xlCenter
FormatExcelPlage MySeet.Range("W2"), 15, False, True, xlCenter, xlCenter
FormatExcelPlage MySeet.Range("AH2"), 15, False, True, xlCenter, xlCenter

R1 = MySeet.Range(MyRange(2, MyRange.Columns.Count).Address).Row
R2 = MySeet.Range(MyRange(MyRange.Rows.Count, MyRange.Columns.Count).Address).Row
MySeet.Range("E2").FormulaR1C1 = "=SUBTOTAL(9,R[" & R1 - 2 & "]C:R[" & R2 - 2 & "]C)"
MySeet.Range("O2").FormulaR1C1 = "=SUBTOTAL(9,R[" & R1 - 2 & "]C:R[" & R2 - 2 & "]C)"
MySeet.Range("X2").FormulaR1C1 = "=SUBTOTAL(9,R[" & R1 - 2 & "]C:R[" & R2 - 2 & "]C)"
MySeet.Range("AI2").FormulaR1C1 = "=SUBTOTAL(9,R[" & R1 - 2 & "]C:R[" & R2 - 2 & "]C)"
MySeet.Range("B2").FormulaR1C1 = "=SUM(RC[1]:RC[" & iTotal & "])"

FormatExcelPlage MySeet.Range("E2"), 2, False, True, xlCenter, xlCenter
FormatExcelPlage MySeet.Range("O2"), 2, False, True, xlCenter, xlCenter
FormatExcelPlage MySeet.Range("X2"), 2, False, True, xlCenter, xlCenter
FormatExcelPlage MySeet.Range("AI2"), 2, False, True, xlCenter, xlCenter
FormatExcelPlage MySeet.Range("B2"), 2, False, True, xlCenter, xlCenter


MiseEnPage MySeet, MySeet.Range("A5").CurrentRegion, "Affaire : " & RsEntetePage!CleAc & Chr(10) & _
 _
     "Pièce : " & RsEntetePage!Piece & Chr(10) & "Liste : " & RsEntetePage!Liste, vbCrLf & "Câblage : " & Replace("" & RsEntetePage!Ensemble, vbCrLf, " ") & Chr(10) & "Equipement : " & Replace("" & RsEntetePage!Equipement, vbCrLf, " ") & Chr(10) & _
     "" _
    , "ENCELADE" & Chr(10) & "Client : " & RsEntetePage!Client & Chr(10) & Format(Date, "dd-mmm-yyyy"), "", "&P/&N", "", 51, "B6", True, 2, True

    MaJEncadreXls MySeet.Range("A5").CurrentRegion, xlThin, xlThin, xlHairline, xlHairline

Sql = "SELECT RqCartouche.* "
Sql = Sql & "FROM RqCartouche "
 Sql = Sql & "WHERE T_indiceProjet.Id=" & Id_IndiceProjet & ";"
Set Rs = Con.OpenRecordSet(Sql)

'NomenclatureHabillage Id_IndiceProjet, MyWorkbook
 AfficheErreur PathPl, EnteteCartouche("" & Rs.Fields("Ensemble"), "" & Rs.Fields("PL_Indice"), "" & Rs.Fields("Li"))
 
MyWordSaveAs PathPl, Save
 If IdFils <> 0 Then
        Sql = "SELECT RqCartouche.* "
        Sql = Sql & "FROM RqCartouche "
        Sql = Sql & "WHERE T_indiceProjet.Id=" & IdFils & ";"
        Set Rs2 = Con.OpenRecordSet(Sql)
         PathPl2 = PathArchive(TableauPath("PathArchiveAutocad"), "" & Rs2!Client, "" & Rs2!CleAc, "" & Rs2!Pieces, "Li", Rs2.Fields("Li"), IdFils, Rs2.Fields("PI_Indice"), Rs2.Fields("LI_Indice"), Rs2!Version)
       Racourci "" & PathPl2 & "_ETIQUETTE", PathPl & "_ETIQUETTE", "DOC"
       Racourci "" & PathPl2 & "_ETIQUETTE_MARQUAGE", PathPl & "_ETIQUETTE_MARQUAGE", "DOC"
    End If
Nomenclature2 = True
insertExelAccess MySeet, "T_Nomenclature", 5, Id_IndiceProjet
End Function

Function Nomenclature(Id_IndiceProjet As Long, Optional PathPl As String, Optional Save As Boolean = True) As Boolean
'\\10.30.0.5\production\Cablage-production\AutoCable\Access\eBoutique
Dim TxtWhere As String
 Dim TbAlve As New Collection
Dim ConConnecteur As New Ado
Dim Sql As String
Dim Rs As Recordset
Dim RsWhere As Recordset
Dim Rs2 As Recordset
Dim NumFieldsConnecteur As Long
Dim NumFieldsConnecteurQts As Long
Dim NumFieldsFournisseur As Long
Dim NumFieldsBouchon As Long
Dim NumFieldsCapot As Long
Dim NumFieldsVerou As Long
Dim NumFieldsJoint As Long
Dim NumFieldsAlveole As Long
Dim DbOk As Boolean
Dim Id As Long
Dim MyEtiquette As New ClsEtiqette
Dim MyRange As Range
Dim PathModelWord As String
Dim PIE(1) As String
Dim Ensemble(1) As String
Dim I As Long
Dim RsComposant As Recordset
Dim RsFils As Recordset
Dim AlveolOk As Boolean
Dim iTotal As Long
Dim MySeet As Worksheet
Set MySeet = IsertSheet(MyWorkbook, "Appro Connectique", True)
'MySeet.Application'.Visible = True

ConConnecteur.TYPEBASE = ADO_TYPEBASE
ConConnecteur.SERVER = ADO_SERVER
ConConnecteur.User = ADO_User
ConConnecteur.PassWord = ADO_PassWord
ConConnecteur.BASE = DbCatalogue

ConConnecteur.BASE = DefinirChemienComplet(TableauPath.Item("PathServer"), TableauPath("Eb_CONNECTEURS"))
DbOk = ConConnecteur.OpenConnetion
DeleteRow MySeet, True

Sql = "SELECT [PI] & '_' & [PI_Indice] AS PIE, T_indiceProjet.Ensemble "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Id_IndiceProjet & ";"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = False Then
    PIE(0) = "PI": PIE(1) = "" & Rs!PIE
    Ensemble(0) = "Ensemble": Ensemble(1) = "" & Rs!Ensemble
End If

Sql = "SELECT Rq_Compte_Connecteur_IdPices.CONNECTEUR, Rq_Compte_Connecteur_IdPices.OPTION, Rq_Compte_Connecteur_IdPices.[Qté]"
Sql = Sql & "FROM Rq_Compte_Connecteur_IdPices "
Sql = Sql & "WHERE Rq_Compte_Connecteur_IdPices.CONNECTEUR<>'NEANT' "
Sql = Sql & "AND Rq_Compte_Connecteur_IdPices.Id_IndiceProjet=" & Id_IndiceProjet & " " '94 "
Sql = Sql & "ORDER BY Rq_Compte_Connecteur_IdPices.CONNECTEUR;"
Set Rs = Con.OpenRecordSet(Sql)
NumFieldsConnecteur = Rs.Fields.Count + 1
    If Rs.EOF = True Then Exit Function
    
'Set MyWord = New Word.Application



PathModelWord = TableauPath.Item("PathModelWord")
PathModelWord = DefinirChemienComplet(TableauPath.Item("PathServer"), PathModelWord)
'         If Left(PathModelWord, 2) <> "\\" And Left(PathModelWord, 1) = "\" Then PathModelWord = TableauPath.Item("PathServer") & PathModelWord
'          If Right(PathModelWord, 2) = "\\" Then PathModelWord = Mid(PathModelWord, 1, Len(PathModelWord) - 1)
Set MyWordDoc = WordNewDoc(PathModelWord)
          
' MyWordDoc.Application'.Visible = True
PathModelWord = TableauPath.Item("PathModelWordMarc")
PathModelWord = DefinirChemienComplet(TableauPath.Item("PathServer"), PathModelWord)
'         If Left(PathModelWord, 2) <> "\\" And Left(PathModelWord, 1) = "\" Then PathModelWord = TableauPath.Item("PathServer") & PathModelWord
'          If Right(PathModelWord, 2) = "\\" Then PathModelWord = Mid(PathModelWord, 1, Len(PathModelWord) - 1)
Set MyWordDoc2 = WordNewDoc(PathModelWord)
          
'MyExcel'.Visible = True
    
While Rs.EOF = False
NbLigne = NbLigne + 1
Rs.MoveNext
Wend
Rs.Requery
'MyWorkbook.Application'.Visible = True
For I = 0 To Rs.Fields.Count - 1
'MySeet.Application'.Visible = True
    MySeet.Cells(5, I + 1) = Rs(I).Name
Next
Row = 5
 FormBarGrah.ProgressBar1.Value = 0
 If NbLigne = 0 Then NbLigne = 1
 FormBarGrah.ProgressBar1.Max = NbLigne
 FormBarGrah.ProgressBar1Caption.Caption = " Exporter liste des Appros :"
DoEvents
While Rs.EOF = False
Row = Row + 1
  IncremanteBarGrah FormBarGrah
   Connecteur = "" & Rs(0).Value
   If InStr(1, "" & Connecteur, "§") <> 0 Then
    Connecteur = Split(Connecteur, "§")
    Connecteur = Connecteur(0)
   End If
    MySeet.Cells(Row, 1) = Connecteur
   MySeet.Cells(Row, 2) = "" & Rs(1)
   MySeet.Cells(Row, 3) = "" & Rs(2)
'      MySeet.Cells(Row, 4) = "" & Rs(3)
'      MySeet.Cells(Row, 5).FormulaR1C1 = "" & Rs(4)


    ChargeLienObjet "RefConnecteur", DefinirChemienComplet(TableauPath.Item("PathServer"), TableauPath("Eb_CONNECTEURS")), NumFieldsConnecteurQts, NumFieldsConnecteur, "" & Connecteur, MySeet, CLng(Row), Db2:=DefinirChemienComplet(TableauPath.Item("PathServer"), TableauPath("Eb_CONNECTIQUE"))
      Sql = "SELECT Connecteurs.CODE_APP, Connecteurs.DESIGNATION "
    Sql = Sql & "FROM Connecteurs "
    Sql = Sql & "WHERE Connecteurs.Id_IndiceProjet=" & Id_IndiceProjet & " "
    Sql = Sql & "AND Connecteurs.CONNECTEUR='" & Connecteur & "' "
    If Trim(MyReplace("" & Rs!Option)) = "" Then
        Sql = Sql & "AND Connecteurs.[O/N]=False AND Connecteurs.OPTION is null ;"
    Else
        Sql = Sql & "AND Connecteurs.[O/N]=False AND Connecteurs.OPTION='" & MyReplace("" & Rs!Option) & "' ;"
    End If
     Set Rs2 = Con.OpenRecordSet(Sql)
        NumFieldsConnecteurApp = NumFieldsConnecteurQts + Rs2.Fields.Count
     MyEtiquette.PrpareEtiqet Rs2.GetString, ""
     MyEtiquette.RenseigneChamp "Connecteur", "" & Connecteur
     MyEtiquette.RenseigneChamp "" & PIE(0), "" & PIE(1)
     MyEtiquette.RenseigneChamp "" & Ensemble(0), "" & Ensemble(1)
    For I = 0 To Rs2.Fields.Count - 1
    
        MySeet.Cells(5, I + NumFieldsConnecteurQts) = Rs2(I).Name
    
    Next
    Rs2.Requery
    If Rs2.EOF = True Then
  
        FunError 8, "" & Connecteur, "" & vbCrLf & Err.Description
    End If
    
    While Rs2.EOF = False
    
            For I = 0 To Rs2.Fields.Count - 1
        MySeet.Cells(Row, I + NumFieldsConnecteurQts) = MySeet.Cells(Row, I + NumFieldsConnecteurQts) & Chr(10) & Replace("" & Rs2(I), Chr(13), "")
         
        Next
             Rs2.MoveNext
        Wend
   If DbOk = True Then
    Sql = "SELECT Rq_Fournisseur.* "
    Sql = Sql & "FROM Rq_Fournisseur "
    Sql = Sql & "WHERE Rq_Fournisseur.[Ref Connecteur]= '" & Connecteur & "';"
    
    Sql = "SELECT lst6.CatName AS Couleur, con_contacts.mem1 AS [Lib Connecteur],  "
    Sql = Sql & "lst9.CatName AS Fournisseur, con_contacts.txt3 AS [Ref Four] "
    Sql = Sql & "FROM (con_contacts INNER JOIN lst9 ON con_contacts.lst9 = lst9.CatID)  "
    Sql = Sql & "LEFT JOIN lst6 ON con_contacts.lst6 = lst6.CatID "
    Sql = Sql & "WHERE lst9.CatName<>'(Sélectionner)'  "
    Sql = Sql & "AND con_contacts.txt1= '" & Connecteur & "';"


    Set Rs2 = ConConnecteur.OpenRecordSet(Sql)
    NumFieldsFournisseur = NumFieldsConnecteurApp + Rs2.Fields.Count
    For I = 0 To Rs2.Fields.Count - 1
        MySeet.Cells(5, I + NumFieldsConnecteurApp) = Rs2(I).Name
    
    Next
    If Rs2.EOF = True Then
  
        FunError 8, "" & Connecteur, "" & vbCrLf & Err.Description
    End If
    While Rs2.EOF = False
    
            For I = 0 To Rs2.Fields.Count - 1
        MySeet.Cells(Row, I + NumFieldsConnecteurApp) = MySeet.Cells(Row, I + NumFieldsConnecteurApp) & Chr(10) & Replace("" & Rs2(I), Chr(13), "")
         
        Next
             Rs2.MoveNext
        Wend
       
        
         
           ChargeLienObjet "RefBouchon", DefinirChemienComplet(TableauPath.Item("PathServer"), TableauPath("Eb_JOINTS")), NumFieldsBouchon, NumFieldsFournisseur, "" & Connecteur, MySeet, CLng(Row)
           ChargeLienObjet "RefCapot", DefinirChemienComplet(TableauPath.Item("PathServer"), TableauPath("Eb_CAPOTS")), NumFieldsCapot, NumFieldsBouchon, "" & Connecteur, MySeet, CLng(Row)
           
'        Sql = "SELECT Rq_Capot.* "
'    Sql = Sql & "FROM Rq_Capot "
'    Sql = Sql & "WHERE Rq_Capot.Référence= '" & Connecteur & "';"
'    Set Rs2 = ConConnecteur.OpenRecordSet(Sql)
  
'    NumFieldsCapot = NumFieldsBouchon + Rs2.Fields.Count - 1
'    For I = 1 To Rs2.Fields.Count - 1
'          MySeet.Cells(5, I + NumFieldsBouchon) = Rs2(I).Name
'
'    Next
'     While Rs2.EOF = False
'
'            For I = 0 To Rs2.Fields.Count - 1
'          MySeet.Cells(Row, I + NumFieldsBouchon) = MySeet.Cells(Row, I + NumFieldsBouchon) & Chr(10) & Replace("" & Rs2(I), Chr(13), "")
'
'        Next
'             Rs2.MoveNext
'        Wend
        
       
'        Sql = "SELECT Rq_Bouchon.* "
'    Sql = Sql & "FROM Rq_Bouchon "
''    Sql = Sql & "WHERE Rq_Bouchon.Référence= '" & Connecteur & "';"
'    Set Rs2 = ConConnecteur.OpenRecordSet(Sql)
'
'
'    NumFieldsBouchon = NumFieldsFournisseur + Rs2.Fields.Count - 1
'    For I = 1 To Rs2.Fields.Count - 1
'          MySeet.Cells(5, I + NumFieldsFournisseur) = Replace(Replace(Replace(Rs2(I).Name, "Qté", "Bouchon Qté"), "Prix U", "Bouchon Prix U"), "Prix Total", "Bouchon Prix Total")
'
'         If Rs2(I).Name = "Qté" Or Rs2(I).Name = "Prix U" Then
'
'            MySeet.Cells(Row, I + NumFieldsFournisseur) = 0
'         End If
'         If Rs2(I).Name = "Prix Total" Then
'            MySeet.Cells(Row, I + NumFieldsFournisseur).FormulaR1C1 = "=(RC[-1]*RC[-2])"
'         End If
'    Next
'     While Rs2.EOF = False
'
'            For I = 1 To Rs2.Fields.Count - 1
'
'
'                             If Rs2(I).Name = "Qté" Or Rs2(I).Name = "Prix U" Or Rs2(I).Name = "Prix Total" Then
'
'                              Else
'                                MySeet.Cells(Row, I + NumFieldsFournisseur) = MySeet.Cells(Row, I + NumFieldsFournisseur) & Chr(10) & Replace("" & Rs2(I), Chr(13), "")
'                  End If
'                  Next
'             Rs2.MoveNext
'        Wend
        
           
'        Sql = "SELECT Rq_Capot.* "
'    Sql = Sql & "FROM Rq_Capot "
'    Sql = Sql & "WHERE Rq_Capot.Référence= '" & Connecteur & "';"
'    Set Rs2 = ConConnecteur.OpenRecordSet(Sql)
'
'    NumFieldsCapot = NumFieldsBouchon + Rs2.Fields.Count - 1
'    For I = 1 To Rs2.Fields.Count - 1
'          MySeet.Cells(5, I + NumFieldsBouchon) = Rs2(I).Name
'
'    Next
'     While Rs2.EOF = False
'
'            For I = 0 To Rs2.Fields.Count - 1
'          MySeet.Cells(Row, I + NumFieldsBouchon) = MySeet.Cells(Row, I + NumFieldsBouchon) & Chr(10) & Replace("" & Rs2(I), Chr(13), "")
'
'        Next
'             Rs2.MoveNext
'        Wend

         ChargeLienObjet "RefVerrou", DefinirChemienComplet(TableauPath.Item("PathServer"), TableauPath("Eb_CAPOTS")), NumFieldsVerou, NumFieldsCapot, "" & Connecteur, MySeet, CLng(Row)
'      Sql = "SELECT Rq_Verou.* "
'    Sql = Sql & "FROM Rq_Verou "
'    Sql = Sql & "WHERE Rq_Verou.Référence= '" & Connecteur & "';"
'    Set Rs2 = ConConnecteur.OpenRecordSet(Sql)
'
'
'   NumFieldsVerou = NumFieldsCapot + Rs2.Fields.Count - 1
'    For I = 1 To Rs2.Fields.Count - 1
'          MySeet.Cells(5, I + NumFieldsCapot) = Rs2(I).Name
'
'    Next
'     While Rs2.EOF = False
'
'            For I = 0 To Rs2.Fields.Count - 1
'          MySeet.Cells(Row, I + NumFieldsCapot) = MySeet.Cells(Row, I + NumFieldsCapot) & Chr(10) & Replace("" & Rs2(I), Chr(13), "")
'
'        Next
'             Rs2.MoveNext
'        Wend

 ChargeLienObjet "RefJoint", DefinirChemienComplet(TableauPath.Item("PathServer"), TableauPath("Eb_JOINTS")), NumFieldsJoint, NumFieldsVerou, "" & Connecteur, MySeet, CLng(Row)
'         Sql = "SELECT Rq_Joint.* "
'    Sql = Sql & "FROM Rq_Joint "
'    Sql = Sql & "WHERE Rq_Joint.Référence= '" & Connecteur & "';"
'    Set Rs2 = ConConnecteur.OpenRecordSet(Sql)
'
'  NumFieldsJoint = NumFieldsVerou + Rs2.Fields.Count - 1
'    For I = 1 To Rs2.Fields.Count - 1
'          MySeet.Cells(5, I + NumFieldsVerou) = Replace(Replace(Replace(Rs2(I).Name, "Qté", "Joint Qté"), "Prix U", "Joint Prix U"), "Prix Total", "Joint Prix Total")
'           If Rs2(I).Name = "Qté" Or Rs2(I).Name = "Prix U" Then
'            MySeet.Cells(Row, I + NumFieldsVerou) = 0
'           End If
'           If Rs2(I).Name = "Prix Total" Then
'           MySeet.Cells(Row, I + NumFieldsVerou).FormulaR1C1 = "=(RC[-1]*RC[-2])"
'           End If
'
'    Next
'     While Rs2.EOF = False
'
'            For I = 1 To Rs2.Fields.Count - 1
'
'
'                 If Rs2(I).Name = "Qté" Or Rs2(I).Name = "Prix U" Or Rs2(I).Name = "Prix Total" Then
'
'                Else
'                MySeet.Cells(Row, I + NumFieldsVerou) = MySeet.Cells(Row, I + NumFieldsVerou) & Chr(10) & Replace("" & Rs2(I), Chr(13), "")
'                End If
'
'          MyEtiquette.RenseigneChamp "" & Rs2(I).Name, "" & Rs2(I).Value
'        Next
'             Rs2.MoveNext
''        End If
'Wend
DoEvents
'ChargeLienObjet "Refclip", "\\10.30.0.5\production\Cablage-production\AutoCable\Access\eBoutique\Encelade_CONNECTIQUE.mdb", NumFieldsJoint, NumFieldsVerou, "" & Connecteur, MySeet, CLng(Row)
AlveolOk = False
'If "7703297847" = Connecteur Then
'MsgBox ""
'End If

Sql = "DELETE TempFillesSection.* FROM TempFillesSection "
Sql = Sql & "WHERE TempFillesSection.Id_IndiceProjet=" & Id_IndiceProjet & ";"
'ConConnecteur.ExecuteSql
Con.Execute Sql

 Sql = "SELECT TempFillesSection.* FROM TempFillesSection "
    Sql = Sql & "WHERE TempFillesSection.Id_IndiceProjet=" & Id_IndiceProjet & ";"

  Set RsFils = Con.OpenRecordSet(Sql)
  
  
  

    Sql = "SELECT Ligne_Tableau_fils.Id_IndiceProjet, Ligne_Tableau_fils.SECT, Ligne_Tableau_fils.VOI AS Vois, Ligne_Tableau_fils.APP AS Cod_App,  "
    Sql = Sql & "Connecteurs.CONNECTEUR "
    Sql = Sql & "FROM Ligne_Tableau_fils INNER JOIN Connecteurs ON (Ligne_Tableau_fils.APP = Connecteurs.CODE_APP)  "
    Sql = Sql & "AND (Ligne_Tableau_fils.Id_IndiceProjet = Connecteurs.Id_IndiceProjet) "
    Sql = Sql & "GROUP BY Ligne_Tableau_fils.Id_IndiceProjet, Ligne_Tableau_fils.SECT, Ligne_Tableau_fils.VOI,  "
    Sql = Sql & "Ligne_Tableau_fils.APP, Connecteurs.CONNECTEUR "
    Sql = Sql & "Having Ligne_Tableau_fils.Id_IndiceProjet = " & Id_IndiceProjet & " And Ligne_Tableau_fils.VOI Is Not Null  "
    Sql = Sql & "And Ligne_Tableau_fils.App Is Not Null And Connecteurs.Connecteur =  '" & Connecteur & "' "
    Sql = Sql & "ORDER BY Ligne_Tableau_fils.APP;"
DoEvents
Set Rs2 = Con.OpenRecordSet(Sql)
DoEvents
While Rs2.EOF = False
RsFils.AddNew
RsFils!Id_IndiceProjet = Rs2!Id_IndiceProjet
RsFils!SECT = ConvertTxtAsDouble("" & Rs2!SECT)
RsFils!VOI = "" & Rs2!VOIS
RsFils!App = "" & Rs2!Cod_App
RsFils!Connecteur = "" & Rs2!Connecteur
RsFils.Update
DoEvents
    Rs2.MoveNext
Wend
  Set Rs2 = Con.CloseRecordSet(Rs2)

        Sql = "SELECT Ligne_Tableau_fils.Id_IndiceProjet, Ligne_Tableau_fils.SECT, Ligne_Tableau_fils.VOI2 AS Vois, "
    Sql = Sql & "Ligne_Tableau_fils.APP2 AS Cod_App, Connecteurs.CONNECTEUR "
    Sql = Sql & "FROM Ligne_Tableau_fils INNER JOIN Connecteurs ON (Ligne_Tableau_fils.APP2 = Connecteurs.CODE_APP) "
    Sql = Sql & "AND (Ligne_Tableau_fils.Id_IndiceProjet = Connecteurs.Id_IndiceProjet) "
    Sql = Sql & "GROUP BY Ligne_Tableau_fils.Id_IndiceProjet, Ligne_Tableau_fils.SECT, Ligne_Tableau_fils.VOI2,  "
    Sql = Sql & "Ligne_Tableau_fils.APP2, Connecteurs.CONNECTEUR "
    Sql = Sql & "Having Ligne_Tableau_fils.Id_IndiceProjet = " & Id_IndiceProjet & " And Ligne_Tableau_fils.VOI2 Is Not Null "
    Sql = Sql & "And Ligne_Tableau_fils.APP2 Is Not Null And Connecteurs.Connecteur =  '" & Connecteur & "' "
    Sql = Sql & "ORDER BY Ligne_Tableau_fils.APP2;"

DoEvents
Set Rs2 = Con.OpenRecordSet(Sql)
DoEvents
While Rs2.EOF = False
RsFils.AddNew
RsFils!Id_IndiceProjet = Rs2!Id_IndiceProjet
RsFils!SECT = ConvertTxtAsDouble("" & Rs2!SECT)
RsFils!VOI = "" & Rs2!VOIS
RsFils!App = "" & Rs2!Cod_App
RsFils!Connecteur = "" & Rs2!Connecteur
RsFils.Update
DoEvents
    Rs2.MoveNext
Wend
 DoEvents
'    Sql = "INSERT INTO TempFillesSection ( Id_IndiceProjet, VOI, SECT, APP ) IN '" & ConConnecteur.RetournDbName("MDB") & "' "
'    Sql = Sql & "SELECT TempFillesSection.Id_IndiceProjet, TempFillesSection.VOI, Sum(TempFillesSection.SECT) AS SommeDeSECT, TempFillesSection.APP "
'    Sql = Sql & "FROM TempFillesSection "
'    Sql = Sql & "GROUP BY TempFillesSection.Id_IndiceProjet, TempFillesSection.VOI, TempFillesSection.APP "
'    Sql = Sql & "HAVING TempFillesSection.Id_IndiceProjet=" & Id_IndiceProjet & ";"
'    DoEvents
'Con.Execute Sql
' MySeconde 2
'
'      sql = "SELECT Rq_Alveole.* "
'    sql = sql & "FROM Rq_Alveole "
'    sql = sql & "WHERE Rq_Alveole.Référence= '" & Connecteur & "';"
'    Set Rs2 = ConConnecteur.OpenRecordSet(sql)
'
DoEvents
    
    
    Sql = "SELECT Rq_Alveole.Référence, Rq_Alveole.[Nb Alvé], Rq_Alveole.Voie, Rq_Alveole.Famille, Rq_Alveole.[Famille Lib],  "
    Sql = Sql & "Rq_Alveole.[Alvé Réf], Rq_Alveole.Qté, Rq_Alveole.[Prix U], Rq_Alveole.[Prix Total], Rq_Alveole.[Alvé Réf Fourr],  "
    Sql = Sql & "Rq_Alveole.[Alvéole Mini en mm2], Rq_Alveole.[Alvéole Maxi en mm2] "
    Sql = Sql & "FROM Rq_Alveole "
    Sql = Sql & "WHERE Rq_Alveole.Id_IndiceProjet=" & Id_IndiceProjet & " AND Rq_Alveole.Référence='" & Connecteur & "' "
    Sql = Sql & "GROUP BY Rq_Alveole.Référence, Rq_Alveole.[Nb Alvé], Rq_Alveole.Voie, Rq_Alveole.Famille, Rq_Alveole.[Famille Lib],  "
    Sql = Sql & "Rq_Alveole.[Alvé Réf], Rq_Alveole.Qté, Rq_Alveole.[Prix U], Rq_Alveole.[Prix Total], Rq_Alveole.[Alvé Réf Fourr],  "
    Sql = Sql & "Rq_Alveole.[Alvéole Mini en mm2], Rq_Alveole.[Alvéole Maxi en mm2], Rq_Alveole.Id_IndiceProjet, Rq_Alveole.Référence; "
    
    Sql = "SELECT T_Alve_Eboutique.[Alvé Réf],T_Alve_Eboutique.[Famille Lib], T_Alve_Eboutique.Mini, T_Alve_Eboutique.Maxi,  "
    Sql = Sql & "T_Alve_Eboutique.[Alvé Réf Fourr],0 as Qté, T_Alve_Eboutique.[Prix u] "
    Sql = Sql & "FROM TempFillesSection, T_LientConnecteur INNER JOIN T_Alve_Eboutique ON T_LientConnecteur.Refclip = T_Alve_Eboutique.[Alvé Réf] "
    Sql = Sql & "GROUP BY T_Alve_Eboutique.[Alvé Réf], T_Alve_Eboutique.Mini, T_Alve_Eboutique.Maxi,  "
    Sql = Sql & "T_Alve_Eboutique.[Famille Lib], T_Alve_Eboutique.[Alvé Réf Fourr], T_Alve_Eboutique.[Prix u],  "
    Sql = Sql & "T_LientConnecteur.RefConnecteur, T_Alve_Eboutique.Mini, T_Alve_Eboutique.Maxi, TempFillesSection.APP, TempFillesSection.SECT "
    Sql = Sql & "HAVING T_LientConnecteur.RefConnecteur='" & Connecteur & "' AND T_Alve_Eboutique.Mini<=[SECT]  "
    Sql = Sql & "AND T_Alve_Eboutique.Maxi>=[SECT]  "
    Sql = Sql & "AND TempFillesSection.Id_IndiceProjet=" & Id_IndiceProjet & ";"

    Sql = "SELECT distinct  T_Alve_Eboutique.[Alvé Réf], T_Alve_Eboutique.[Famille Lib], T_Alve_Eboutique.[Alvé Réf Fourr],  "
    Sql = Sql & "T_Alve_Eboutique.Mini as [Alvéole Mini en mm2], T_Alve_Eboutique.Maxi as [Alvéole Maxi en mm2], 0 AS Qté, 0 AS [Alvé Prix u], 0 AS [Prix Total] "
    Sql = Sql & "FROM TempFillesSection, T_LientConnecteur INNER JOIN T_Alve_Eboutique  "
    Sql = Sql & "ON T_LientConnecteur.Refclip = T_Alve_Eboutique.[Alvé Réf] "
    Sql = Sql & "Where T_Alve_Eboutique.Mini <= [SECT] "
    Sql = Sql & "And T_Alve_Eboutique.Maxi >= [SECT] "
    Sql = Sql & "And T_LientConnecteur.RefConnecteur = '" & Connecteur & "' "
    Sql = Sql & "And TempFillesSection.Id_IndiceProjet = " & Id_IndiceProjet & " "
    Sql = Sql & "GROUP BY TempFillesSection.APP, TempFillesSection.VOI,T_Alve_Eboutique.[Alvé Réf], T_Alve_Eboutique.[Famille Lib],  "
    Sql = Sql & "T_Alve_Eboutique.[Alvé Réf Fourr], T_Alve_Eboutique.[Alvé Fornisseur],  "
    Sql = Sql & "T_Alve_Eboutique.Mini, T_Alve_Eboutique.Maxi, 0, 0, 0, TempFillesSection.SECT,  "
    Sql = Sql & "T_LientConnecteur.RefConnecteur, TempFillesSection.Id_IndiceProjet;"

'Modifier par RD
 Sql = "SELECT DISTINCT T_Lien_Connecteur_Clip.Refclip AS [Alvé Réf],lst21.CatName AS [Famille Lib], con_contacts.txt3 AS [Alvé Réf Fourr],  "
    Sql = Sql & " lst22.CatName AS [Alvéole Mini en mm2], lst23.CatName AS  "
    Sql = Sql & "[Alvéole Maxi en mm2], 0 AS Qté, 0 AS [Prix U], 0 AS [Prix Total] "
    Sql = Sql & "FROM ((((TempFillesSection INNER JOIN T_Lien_Connecteur_Clip ON TempFillesSection.CONNECTEUR =  "
    Sql = Sql & "T_Lien_Connecteur_Clip.RefConnecteur) INNER JOIN con_contacts ON T_Lien_Connecteur_Clip.Refclip =  "
    Sql = Sql & "con_contacts.txt1) LEFT JOIN lst21 ON con_contacts.lst21 = lst21.CatID) INNER JOIN lst22  "
    Sql = Sql & "ON con_contacts.lst22 = lst22.CatID) INNER JOIN lst23 ON con_contacts.lst23 = lst23.CatID "
    Sql = Sql & "WHERE T_Lien_Connecteur_Clip.RefConnecteur= '" & Connecteur & "' "
    Sql = Sql & "AND TempFillesSection.Id_IndiceProjet= " & Id_IndiceProjet & " "
    Sql = Sql & "GROUP BY T_Lien_Connecteur_Clip.Refclip,  lst21.CatName,con_contacts.txt3, lst22.CatName,  "
    Sql = Sql & "lst23.CatName, 0, 0, 0, TempFillesSection.APP, TempFillesSection.VOI,  "
    Sql = Sql & "T_Lien_Connecteur_Clip.RefConnecteur, TempFillesSection.Id_IndiceProjet "
    Sql = Sql & "HAVING (((lst22.CatName)<=Sum([SECT])) AND ((lst23.CatName)>=Sum([SECT])));"

    
    DoEvents
 Set Rs2 = Con.OpenRecordSet(Sql)
' ChargeLienObjet "RefAlve", DefinirChemienComplet(TableauPath.Item("PathServer"), TableauPath("Eb_CONNECTIQUE")), NumFieldsAlveole, NumFieldsJoint, "" & Connecteur, MySeet, CLng(Row)
 DoEvents
'    MyWorkbook.Application'.Visible = True
  
   Set TableauAlve = Nothing
   
 DoEvents
   If Rs2.EOF = False Then
        AlveolOk = True
        TableauAlve = Rs2.GetRows
        DoEvents
        ReDim tableauAlve2(UBound(TableauAlve), 1)
        DoEvents
        I = 0
        For I = 0 To UBound(TableauAlve, 2)
            On Error Resume Next
            a = ""
            a = TbAlve(TableauAlve(3, I))
            If Err <> 0 Then
                Err.Clear
                TbAlve.Add I, TableauAlve(0, I)
            End If
            tableauAlve2(TbAlve(TableauAlve(3, I)), 0) = "" & TableauAlve(1, I) & ": "
             tableauAlve2(TbAlve(TableauAlve(3, I)), 1) = tableauAlve2(TbAlve(TableauAlve(3, I)), 1) & "" & TableauAlve(2, I) & "(_____), "
             DoEvents
        Next
        Txt = ""
        For I = LBound(tableauAlve2) To UBound(tableauAlve2)
            If Trim("" & tableauAlve2(I, 1)) <> "" Then
           Txt = Txt & tableauAlve2(I, 0) & tableauAlve2(I, 1) & ";"
           
            Debug.Print Txt
            
            End If
            DoEvents
        Next
    Else
        Txt = ""
        ReDim tableauAlve2(0, 1)
     End If

Txt = Replace(Txt, ",;", "; ")
Txt = Replace(Txt, ", ;", "; ")
Debug.Print Txt
'txt = Replace(txt, ":", "")
MyEtiquette.RenseigneChamp "Famille", "" & Txt
' MySeet.Application'.Visible = True
'  ReDim tableauAlve(1, 1) Famille
  Dim T_Alve() As String
' ReDim T_Alve(Bound(tableauAlve), 1)
  Rs2.Requery
  NumFieldsAlveole = NumFieldsJoint + Rs2.Fields.Count - 1
    For I = 0 To Rs2.Fields.Count - 1
        If iTotal < I + NumFieldsJoint Then iTotal = I + NumFieldsJoint
          MySeet.Cells(5, I + NumFieldsJoint) = Replace(Replace(Replace(Rs2(I).Name, "Qté", "Alvé Qté"), "Prix U", "Alvé Prix U"), "Prix Total", "Alvé Prix Total")
           If Rs2(I).Name = "Qté" Or Rs2(I).Name = "Prix U" Then
           MySeet.Cells(Row, I + NumFieldsJoint) = 0
         End If
         If Rs2(I).Name = "Prix Total" Then
            MySeet.Cells(Row, I + NumFieldsJoint) = "=(RC[-1]*RC[-2])"
         End If
         
    Next
     While Rs2.EOF = False
   
            For I = 0 To Rs2.Fields.Count - 1
            
                If Rs2(I).Name = "Qté" Or Rs2(I).Name = "Prix U" Or Rs2(I).Name = "Prix Total" Or Rs2(I).Name = "Prix u" Or Rs2(I).Name = "Alvé Prix u" Then
                Else
                    If UCase(Rs2(I).Name) = UCase("Voie") Then
                        MySeet.Cells(Row, I + NumFieldsJoint) = MySeet.Cells(Row, I + NumFieldsJoint) + Val(Replace("" & Rs2(I - 1), Chr(13), ""))
                    Else
                    MySeet.Cells(Row, I + NumFieldsJoint) = MySeet.Cells(Row, I + NumFieldsJoint) & Chr(10) & Replace("" & Rs2(I), Chr(13), "")
                    End If
                End If
        Next
           
             Rs2.MoveNext
        Wend
End If
      DoEvents
'      MyWord.Application'.Visible = True
On Error GoTo Fin
      For I = MyEtiquette.TableMin To MyEtiquette.TableMax
      a = MyEtiquette.RetournEtiquette(I)
        CreerEtiquette MyEtiquette.RetournEtiquette(I)
      Next
Fin:
'      MyWord'.Visible = True
'      CreerEtiquette
    Rs.MoveNext
Wend

Set Rs2 = ConConnecteur.CloseRecordSet(Rs2)
ConConnecteur.CloseConnection
Set Rs = Con.CloseRecordSet(Rs)

Sql = "SELECT  T_indiceProjet.Ensemble, T_indiceProjet.Equipement,[Li] & '_' & [LI_Indice] AS Liste, "
Sql = Sql & "[PI] & '_' & [PI_Indice] AS Piece, T_indiceProjet.Client, T_indiceProjet.CleAc "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Id_IndiceProjet & ";"
Dim RsEntetePage As Recordset
Set RsEntetePage = Con.OpenRecordSet(Sql)
Set MyRange = MySeet.Range("A5").CurrentRegion
'MyWorkbook.Application'.Visible = True
MySeet.Range("A2") = "TOTAL"
MySeet.Range("D2") = "SOUS TOTAL"
MySeet.Range("N2") = "SOUS TOTAL"
MySeet.Range("W2") = "SOUS TOTAL"
MySeet.Range("AH2") = "SOUS TOTAL"

FormatExcelPlage MySeet.Range("A2"), 15, False, True, xlCenter, xlCenter
FormatExcelPlage MySeet.Range("D2"), 15, False, True, xlCenter, xlCenter
FormatExcelPlage MySeet.Range("N2"), 15, False, True, xlCenter, xlCenter
FormatExcelPlage MySeet.Range("W2"), 15, False, True, xlCenter, xlCenter
FormatExcelPlage MySeet.Range("AH2"), 15, False, True, xlCenter, xlCenter

R1 = MySeet.Range(MyRange(2, MyRange.Columns.Count).Address).Row
R2 = MySeet.Range(MyRange(MyRange.Rows.Count, MyRange.Columns.Count).Address).Row
MySeet.Range("E2").FormulaR1C1 = "=SUBTOTAL(9,R[" & R1 - 2 & "]C:R[" & R2 - 2 & "]C)"
MySeet.Range("O2").FormulaR1C1 = "=SUBTOTAL(9,R[" & R1 - 2 & "]C:R[" & R2 - 2 & "]C)"
MySeet.Range("X2").FormulaR1C1 = "=SUBTOTAL(9,R[" & R1 - 2 & "]C:R[" & R2 - 2 & "]C)"
MySeet.Range("AI2").FormulaR1C1 = "=SUBTOTAL(9,R[" & R1 - 2 & "]C:R[" & R2 - 2 & "]C)"
MySeet.Range("B2").FormulaR1C1 = "=SUM(RC[1]:RC[" & iTotal & "])"

FormatExcelPlage MySeet.Range("E2"), 2, False, True, xlCenter, xlCenter
FormatExcelPlage MySeet.Range("O2"), 2, False, True, xlCenter, xlCenter
FormatExcelPlage MySeet.Range("X2"), 2, False, True, xlCenter, xlCenter
FormatExcelPlage MySeet.Range("AI2"), 2, False, True, xlCenter, xlCenter
FormatExcelPlage MySeet.Range("B2"), 2, False, True, xlCenter, xlCenter


MiseEnPage MySeet, MySeet.Range("A5").CurrentRegion, "Affaire : " & RsEntetePage!CleAc & Chr(10) & _
 _
     "Pièce : " & RsEntetePage!Piece & Chr(10) & "Liste : " & RsEntetePage!Liste, vbCrLf & "Câblage : " & Replace("" & RsEntetePage!Ensemble, vbCrLf, " ") & Chr(10) & "Equipement : " & Replace("" & RsEntetePage!Equipement, vbCrLf, " ") & Chr(10) & _
     "" _
    , "ENCELADE" & Chr(10) & "Client : " & RsEntetePage!Client & Chr(10) & Format(Date, "dd-mmm-yyyy"), "", "&P/&N", "", 51, "B6", True, 2, True

    MaJEncadreXls MySeet.Range("A5").CurrentRegion, xlThin, xlThin, xlHairline, xlHairline

Sql = "SELECT RqCartouche.* "
Sql = Sql & "FROM RqCartouche "
 Sql = Sql & "WHERE T_indiceProjet.Id=" & Id_IndiceProjet & ";"
Set Rs = Con.OpenRecordSet(Sql)

'NomenclatureHabillage Id_IndiceProjet, MyWorkbook
 AfficheErreur PathPl, EnteteCartouche("" & Rs.Fields("Ensemble"), "" & Rs.Fields("PL_Indice"), "" & Rs.Fields("Li"))
 
MyWordSaveAs PathPl, Save
 If IdFils <> 0 Then
        Sql = "SELECT RqCartouche.* "
        Sql = Sql & "FROM RqCartouche "
        Sql = Sql & "WHERE T_indiceProjet.Id=" & IdFils & ";"
        Set Rs2 = Con.OpenRecordSet(Sql)
         PathPl2 = PathArchive(TableauPath("PathArchiveAutocad"), "" & Rs2!Client, "" & Rs2!CleAc, "" & Rs2!Pieces, "Li", Rs2.Fields("Li"), IdFils, Rs2.Fields("PI_Indice"), Rs2.Fields("LI_Indice"), Rs2!Version)
       Racourci "" & PathPl2 & "_ETIQUETTE", PathPl & "_ETIQUETTE", "DOC"
       Racourci "" & PathPl2 & "_ETIQUETTE_MARQUAGE", PathPl & "_ETIQUETTE_MARQUAGE", "DOC"
    End If
Nomenclature = True
insertExelAccess MySeet, "T_Nomenclature", 5, Id_IndiceProjet
End Function


Function Nomenclature3(Id_IndiceProjet As Long, Optional PathPl As String, Optional Save As Boolean = True) As Boolean
'\\10.30.0.5\production\Cablage-production\AutoCable\Access\eBoutique
Dim TxtWhere As String
 Dim TbAlve As New Collection
Dim ConConnecteur As New Ado
Dim ConConnectique As New Ado
Dim Sql As String
Dim Rs As Recordset
Dim RsWhere As Recordset
Dim Rs2 As Recordset
Dim NumFieldsConnecteur As Long
Dim NumFieldsConnecteurQts As Long
Dim NumFieldsFournisseur As Long
Dim NumFieldsBouchon As Long
Dim NumFieldsCapot As Long
Dim NumFieldsVerou As Long
Dim NumFieldsJoint As Long
Dim NumFieldsAlveole As Long
Dim DbOk As Boolean
Dim Id As Long
'Dim 'MyEtiquette As New ClsEtiqette
Dim MyRange As Range
Dim PathModelWord As String
Dim PIE(1) As String
Dim Ensemble(1) As String
Dim I As Long
Dim RsComposant As Recordset
Dim RsFils As Recordset
Dim AlveolOk As Boolean
Dim iTotal As Long
Dim MySeet As Worksheet
Set MySeet = IsertSheet(MyWorkbook, "Appro Connectique", True)
Dim NumRepris As Integer
'MySeet.Application'.Visible = True
'
DbOk = True
'DbOk = Con.OpenConnetion(DefinirChemienComplet(TableauPath.Item("PathServer"), TableauPath("Eb_CONNECTEURS"))) '"\\10.30.0.5\production\Cablage-production\AutoCable\Access\eBoutique\le_11_10_05\Encelade_A-CONNECTEURS.mdb")  'DbCatalogue)  '"U:\Librairies\Plans\AutoCâble\Xp\Access\Catalogue Renault.mdb" "\\10.30.0.5\production\Cablage-production\AutoCable\Access\eBoutique\Encelade_CONNECTEURS.mdb") '
'DbOk = ConConnectique.OpenConnetion(DefinirChemienComplet(TableauPath.Item("PathServer"), TableauPath("Eb_CONNECTIQUE"))) '"\\10.30.0.5\production\Cablage-production\AutoCable\Access\eBoutique\le_11_10_05\Encelade_A-CONNECTEURS.mdb")  'DbCatalogue)  '"U:\Librairies\Plans\AutoCâble\Xp\Access\Catalogue Renault.mdb" "\\10.30.0.5\production\Cablage-production\AutoCable\Access\eBoutique\Encelade_CONNECTEURS.mdb") '
DeleteRow MySeet, True

Sql = "SELECT [PI] & '_' & [PI_Indice] AS PIE, T_indiceProjet.Ensemble "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Id_IndiceProjet & ";"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = False Then
    PIE(0) = "PI": PIE(1) = "" & Rs!PIE
    Ensemble(0) = "Ensemble": Ensemble(1) = "" & Rs!Ensemble
End If

Sql = "SELECT Rq_Compte_Connecteur_IdPices.CONNECTEUR, Rq_Compte_Connecteur_IdPices.OPTION, Rq_Compte_Connecteur_IdPices.[Qté]"
Sql = Sql & "FROM Rq_Compte_Connecteur_IdPices "
Sql = Sql & "WHERE Rq_Compte_Connecteur_IdPices.CONNECTEUR<>'NEANT' "
Sql = Sql & "AND Rq_Compte_Connecteur_IdPices.Id_IndiceProjet=" & Id_IndiceProjet & "  " 'and Rq_Compte_Connecteur_IdPices.CONNECTEUR='8200062510' " '94 "
Sql = Sql & "ORDER BY Rq_Compte_Connecteur_IdPices.CONNECTEUR;"
Set Rs = Con.OpenRecordSet(Sql)
NumFieldsConnecteur = Rs.Fields.Count + 1
    If Rs.EOF = True Then Exit Function
    
'Set MyWord = New Word.Application


'
'PathModelWord = TableauPath.Item("PathModelWord")
'PathModelWord = DefinirChemienComplet(TableauPath.Item("PathServer"), PathModelWord)
''         If Left(PathModelWord, 2) <> "\\" And Left(PathModelWord, 1) = "\" Then PathModelWord = TableauPath.Item("PathServer") & PathModelWord
''          If Right(PathModelWord, 2) = "\\" Then PathModelWord = Mid(PathModelWord, 1, Len(PathModelWord) - 1)
'Set MyWordDoc = WordNewDoc(PathModelWord)
'
'' MyWordDoc.Application'.Visible = True
'PathModelWord = TableauPath.Item("PathModelWordMarc")
'PathModelWord = DefinirChemienComplet(TableauPath.Item("PathServer"), PathModelWord)
''         If Left(PathModelWord, 2) <> "\\" And Left(PathModelWord, 1) = "\" Then PathModelWord = TableauPath.Item("PathServer") & PathModelWord
''          If Right(PathModelWord, 2) = "\\" Then PathModelWord = Mid(PathModelWord, 1, Len(PathModelWord) - 1)
'Set MyWordDoc2 = WordNewDoc(PathModelWord)
          
'MyExcel'.Visible = True
    
While Rs.EOF = False
NbLigne = NbLigne + 1
Rs.MoveNext
Wend
Rs.Requery
'MyWorkbook.Application'.Visible = True
For I = 0 To Rs.Fields.Count - 1
'MySeet.Application.Visible = True
    MySeet.Cells(5, I + 1) = Rs(I).Name
Next
Row = 5
 FormBarGrah.ProgressBar1.Value = 0
 If NbLigne = 0 Then NbLigne = 1
 FormBarGrah.ProgressBar1.Max = NbLigne
 FormBarGrah.ProgressBar1Caption.Caption = " Exporter liste des Appros :"
DoEvents
While Rs.EOF = False
Row = Row + 1
  IncremanteBarGrah FormBarGrah
  IncrmentServer FormBarGrah, ""
   Connecteur = "" & Rs(0).Value
   
    Connecteur = Split(Connecteur & "§", "§")
    Connecteur = Connecteur(0)
   
    MySeet.Cells(Row, 1) = Connecteur
   MySeet.Cells(Row, 2) = "" & Rs(1)
   MySeet.Cells(Row, 3) = "" & Rs(2)
'      MySeet.Cells(Row, 4) = "" & Rs(3)
'      MySeet.Cells(Row, 5).FormulaR1C1 = "" & Rs(4)


    ChargeLienObjet "RefConnecteur", TableauPath("Eb_CONNECTEURS"), NumFieldsConnecteurQts, NumFieldsConnecteur, "" & Connecteur, MySeet, CLng(Row), Db2:=TableauPath("Eb_CONNECTIQUE")
      Sql = "SELECT Connecteurs.CODE_APP, Connecteurs.DESIGNATION "
    Sql = Sql & "FROM Connecteurs "
    Sql = Sql & "WHERE Connecteurs.Id_IndiceProjet=" & Id_IndiceProjet & " "
    Sql = Sql & "AND Connecteurs.CONNECTEUR='" & Connecteur & "' "
    If Trim(MyReplace("" & Rs!Option)) = "" Then
        Sql = Sql & "AND Connecteurs.[O/N]=False AND Connecteurs.OPTION is null ;"
    Else
        Sql = Sql & "AND Connecteurs.[O/N]=False AND Connecteurs.OPTION='" & MyReplace("" & Rs!Option) & "' ;"
    End If
     Set Rs2 = Con.OpenRecordSet(Sql)
        NumFieldsConnecteurApp = NumFieldsConnecteurQts + Rs2.Fields.Count
     'MyEtiquette.PrpareEtiqet Rs2.GetString, ""
     ''MyEtiquette.RenseigneChamp "Connecteur", "" & Connecteur
     ''MyEtiquette.RenseigneChamp "" & PIE(0), "" & PIE(1)
     ''MyEtiquette.RenseigneChamp "" & Ensemble(0), "" & Ensemble(1)
    For I = 0 To Rs2.Fields.Count - 1
    
        MySeet.Cells(5, I + NumFieldsConnecteurQts) = Rs2(I).Name
    
    Next
    Rs2.Requery
    If Rs2.EOF = True Then
  
        FunError 8, "" & Connecteur, "" & vbCrLf & Err.Description
    End If
    
    While Rs2.EOF = False
    
            For I = 0 To Rs2.Fields.Count - 1
        MySeet.Cells(Row, I + NumFieldsConnecteurQts) = MySeet.Cells(Row, I + NumFieldsConnecteurQts) & Chr(10) & Replace("" & Rs2(I), Chr(13), "")
         
        Next
             Rs2.MoveNext
        Wend
   If DbOk = True Then
    Sql = "SELECT Rq_Fournisseur.* "
    Sql = Sql & "FROM Rq_Fournisseur IN '"
    Sql = Sql & "WHERE Rq_Fournisseur.[Ref Connecteur]= '" & Connecteur & "';"
    
    Sql = "SELECT lst6.CatName AS Couleur, con_contacts.mem1 AS [Lib Connecteur],  "
    Sql = Sql & "lst9.CatName AS Fournisseur, con_contacts.txt3 AS [Ref Four] "
    Sql = Sql & "FROM (con_contacts INNER JOIN lst9 ON con_contacts.lst9 = lst9.CatID)  "
    Sql = Sql & "LEFT JOIN lst6 ON con_contacts.lst6 = lst6.CatID IN'" & TableauPath("Eb_CONNECTEURS") & "' "
    Sql = Sql & "WHERE lst9.CatName<>'(Sélectionner)'  "
    Sql = Sql & "AND con_contacts.txt1= '" & Connecteur & "';"


    Set Rs2 = Con.OpenRecordSet(Sql)
    If Rs2.EOF = True Then
        Sql = "SELECT '' AS Couleur, con_contacts.mem1 AS [Lib Connecteur],  "
        Sql = Sql & "lst9.CatName AS Fournisseur, con_contacts.txt3 AS [Ref Four] "
        Sql = Sql & "FROM (con_contacts INNER JOIN lst9 ON con_contacts.lst9 = lst9.CatID)  "
        Sql = Sql & "LEFT JOIN lst6 ON con_contacts.lst6 = lst6.CatID IN'" & TableauPath("Eb_CONNECTEURS") & "' "
        Sql = Sql & "WHERE lst9.CatName<>'(Sélectionner)'  "
        Sql = Sql & "AND con_contacts.txt1= '" & Connecteur & "';"
    
    
        Set Rs2 = Con.OpenRecordSet(Sql)

    End If
    NumFieldsFournisseur = NumFieldsConnecteurApp + Rs2.Fields.Count
    For I = 0 To Rs2.Fields.Count - 1
        MySeet.Cells(5, I + NumFieldsConnecteurApp) = Rs2(I).Name
    
    Next
    If Rs2.EOF = True Then
  
        FunError 8, "" & Connecteur, "" & vbCrLf & Err.Description
    End If
 


    While Rs2.EOF = False
    
            For I = 0 To Rs2.Fields.Count - 1
        MySeet.Cells(Row, I + NumFieldsConnecteurApp) = MySeet.Cells(Row, I + NumFieldsConnecteurApp) & Chr(10) & Replace("" & Rs2(I), Chr(13), "")
         
        Next
             Rs2.MoveNext
        Wend
       
  NumRepris = 0
         
           ChargeLienObjet "RefBouchon", TableauPath("Eb_JOINTS"), NumFieldsBouchon, NumFieldsFournisseur, "" & Connecteur, MySeet, CLng(Row)
           ChargeLienObjet "RefCapot", DefinirChemienComplet(TableauPath.Item("PathServer"), TableauPath("Eb_CAPOTS")), NumFieldsCapot, NumFieldsBouchon, "" & Connecteur, MySeet, CLng(Row)
           
'        Sql = "SELECT Rq_Capot.* "
'    Sql = Sql & "FROM Rq_Capot "
'    Sql = Sql & "WHERE Rq_Capot.Référence= '" & Connecteur & "';"
'    Set Rs2 = Con.OpenRecordSet(Sql)
  
'    NumFieldsCapot = NumFieldsBouchon + Rs2.Fields.Count - 1
'    For I = 1 To Rs2.Fields.Count - 1
'          MySeet.Cells(5, I + NumFieldsBouchon) = Rs2(I).Name
'
'    Next
'     While Rs2.EOF = False
'
'            For I = 0 To Rs2.Fields.Count - 1
'          MySeet.Cells(Row, I + NumFieldsBouchon) = MySeet.Cells(Row, I + NumFieldsBouchon) & Chr(10) & Replace("" & Rs2(I), Chr(13), "")
'
'        Next
'             Rs2.MoveNext
'        Wend
        
       
'        Sql = "SELECT Rq_Bouchon.* "
'    Sql = Sql & "FROM Rq_Bouchon "
''    Sql = Sql & "WHERE Rq_Bouchon.Référence= '" & Connecteur & "';"
'    Set Rs2 = Con.OpenRecordSet(Sql)
'
'
'    NumFieldsBouchon = NumFieldsFournisseur + Rs2.Fields.Count - 1
'    For I = 1 To Rs2.Fields.Count - 1
'          MySeet.Cells(5, I + NumFieldsFournisseur) = Replace(Replace(Replace(Rs2(I).Name, "Qté", "Bouchon Qté"), "Prix U", "Bouchon Prix U"), "Prix Total", "Bouchon Prix Total")
'
'         If Rs2(I).Name = "Qté" Or Rs2(I).Name = "Prix U" Then
'
'            MySeet.Cells(Row, I + NumFieldsFournisseur) = 0
'         End If
'         If Rs2(I).Name = "Prix Total" Then
'            MySeet.Cells(Row, I + NumFieldsFournisseur).FormulaR1C1 = "=(RC[-1]*RC[-2])"
'         End If
'    Next
'     While Rs2.EOF = False
'
'            For I = 1 To Rs2.Fields.Count - 1
'
'
'                             If Rs2(I).Name = "Qté" Or Rs2(I).Name = "Prix U" Or Rs2(I).Name = "Prix Total" Then
'
'                              Else
'                                MySeet.Cells(Row, I + NumFieldsFournisseur) = MySeet.Cells(Row, I + NumFieldsFournisseur) & Chr(10) & Replace("" & Rs2(I), Chr(13), "")
'                  End If
'                  Next
'             Rs2.MoveNext
'        Wend
        
           
'        Sql = "SELECT Rq_Capot.* "
'    Sql = Sql & "FROM Rq_Capot "
'    Sql = Sql & "WHERE Rq_Capot.Référence= '" & Connecteur & "';"
'    Set Rs2 = Con.OpenRecordSet(Sql)
'
'    NumFieldsCapot = NumFieldsBouchon + Rs2.Fields.Count - 1
'    For I = 1 To Rs2.Fields.Count - 1
'          MySeet.Cells(5, I + NumFieldsBouchon) = Rs2(I).Name
'
'    Next
'     While Rs2.EOF = False
'
'            For I = 0 To Rs2.Fields.Count - 1
'          MySeet.Cells(Row, I + NumFieldsBouchon) = MySeet.Cells(Row, I + NumFieldsBouchon) & Chr(10) & Replace("" & Rs2(I), Chr(13), "")
'
'        Next
'             Rs2.MoveNext
'        Wend

         ChargeLienObjet "RefVerrou", DefinirChemienComplet(TableauPath.Item("PathServer"), TableauPath("Eb_CAPOTS")), NumFieldsVerou, NumFieldsCapot, "" & Connecteur, MySeet, CLng(Row)
'      Sql = "SELECT Rq_Verou.* "
'    Sql = Sql & "FROM Rq_Verou "
'    Sql = Sql & "WHERE Rq_Verou.Référence= '" & Connecteur & "';"
'    Set Rs2 = Con.OpenRecordSet(Sql)
'
'
'   NumFieldsVerou = NumFieldsCapot + Rs2.Fields.Count - 1
'    For I = 1 To Rs2.Fields.Count - 1
'          MySeet.Cells(5, I + NumFieldsCapot) = Rs2(I).Name
'
'    Next
'     While Rs2.EOF = False
'
'            For I = 0 To Rs2.Fields.Count - 1
'          MySeet.Cells(Row, I + NumFieldsCapot) = MySeet.Cells(Row, I + NumFieldsCapot) & Chr(10) & Replace("" & Rs2(I), Chr(13), "")
'
'        Next
'             Rs2.MoveNext
'        Wend

 ChargeLienObjet "RefJoint", DefinirChemienComplet(TableauPath.Item("PathServer"), TableauPath("Eb_JOINTS")), NumFieldsJoint, NumFieldsVerou, "" & Connecteur, MySeet, CLng(Row)
'         Sql = "SELECT Rq_Joint.* "
'    Sql = Sql & "FROM Rq_Joint "
'    Sql = Sql & "WHERE Rq_Joint.Référence= '" & Connecteur & "';"
'    Set Rs2 = Con.OpenRecordSet(Sql)
'
'  NumFieldsJoint = NumFieldsVerou + Rs2.Fields.Count - 1
'    For I = 1 To Rs2.Fields.Count - 1
'          MySeet.Cells(5, I + NumFieldsVerou) = Replace(Replace(Replace(Rs2(I).Name, "Qté", "Joint Qté"), "Prix U", "Joint Prix U"), "Prix Total", "Joint Prix Total")
'           If Rs2(I).Name = "Qté" Or Rs2(I).Name = "Prix U" Then
'            MySeet.Cells(Row, I + NumFieldsVerou) = 0
'           End If
'           If Rs2(I).Name = "Prix Total" Then
'           MySeet.Cells(Row, I + NumFieldsVerou).FormulaR1C1 = "=(RC[-1]*RC[-2])"
'           End If
'
'    Next
'     While Rs2.EOF = False
'
'            For I = 1 To Rs2.Fields.Count - 1
'
'
'                 If Rs2(I).Name = "Qté" Or Rs2(I).Name = "Prix U" Or Rs2(I).Name = "Prix Total" Then
'
'                Else
'                MySeet.Cells(Row, I + NumFieldsVerou) = MySeet.Cells(Row, I + NumFieldsVerou) & Chr(10) & Replace("" & Rs2(I), Chr(13), "")
'                End If
'
'          ''MyEtiquette.RenseigneChamp "" & Rs2(I).Name, "" & Rs2(I).Value
'        Next
'             Rs2.MoveNext
''        End If
'Wend
DoEvents
'ChargeLienObjet "Refclip", "\\10.30.0.5\production\Cablage-production\AutoCable\Access\eBoutique\Encelade_CONNECTIQUE.mdb", NumFieldsJoint, NumFieldsVerou, "" & Connecteur, MySeet, CLng(Row)
AlveolOk = False
'If "7703297847" = Connecteur Then
'MsgBox ""
'End If

Sql = "DELETE TempFillesSection.* FROM TempFillesSection "
Sql = Sql & "WHERE TempFillesSection.Id_IndiceProjet=" & Id_IndiceProjet & ";"
'Con.Execute Sql
Con.Execute Sql

 Sql = "SELECT TempFillesSection.* FROM TempFillesSection "
    Sql = Sql & "WHERE TempFillesSection.Id_IndiceProjet=" & Id_IndiceProjet & ";"

  Set RsFils = Con.OpenRecordSet(Sql)
  
  
  

    Sql = "SELECT Ligne_Tableau_fils.Id_IndiceProjet, Ligne_Tableau_fils.SECT, Ligne_Tableau_fils.VOI AS Vois, Ligne_Tableau_fils.APP AS Cod_App,  "
    Sql = Sql & "Connecteurs.CONNECTEUR "
    Sql = Sql & "FROM Ligne_Tableau_fils INNER JOIN Connecteurs ON (Ligne_Tableau_fils.APP = Connecteurs.CODE_APP)  "
    Sql = Sql & "AND (Ligne_Tableau_fils.Id_IndiceProjet = Connecteurs.Id_IndiceProjet) "
    Sql = Sql & "GROUP BY Ligne_Tableau_fils.Id_IndiceProjet, Ligne_Tableau_fils.SECT, Ligne_Tableau_fils.VOI,  "
    Sql = Sql & "Ligne_Tableau_fils.APP, Connecteurs.CONNECTEUR "
    Sql = Sql & "Having Ligne_Tableau_fils.Id_IndiceProjet = " & Id_IndiceProjet & " And Ligne_Tableau_fils.VOI Is Not Null  "
    Sql = Sql & "And Ligne_Tableau_fils.App Is Not Null And Connecteurs.Connecteur =  '" & Connecteur & "' "
    Sql = Sql & "ORDER BY Ligne_Tableau_fils.APP;"
DoEvents
Set Rs2 = Con.OpenRecordSet(Sql)
DoEvents
While Rs2.EOF = False
RsFils.AddNew
RsFils!Id_IndiceProjet = Rs2!Id_IndiceProjet
RsFils!SECT = ConvertTxtAsDouble("" & Rs2!SECT)
RsFils!VOI = "" & Rs2!VOIS
RsFils!App = "" & Rs2!Cod_App
RsFils!Connecteur = "" & Rs2!Connecteur
RsFils.Update
DoEvents
    Rs2.MoveNext
Wend
  Set Rs2 = Con.CloseRecordSet(Rs2)

        Sql = "SELECT Ligne_Tableau_fils.Id_IndiceProjet, Ligne_Tableau_fils.SECT, Ligne_Tableau_fils.VOI2 AS Vois, "
    Sql = Sql & "Ligne_Tableau_fils.APP2 AS Cod_App, Connecteurs.CONNECTEUR "
    Sql = Sql & "FROM Ligne_Tableau_fils INNER JOIN Connecteurs ON (Ligne_Tableau_fils.APP2 = Connecteurs.CODE_APP) "
    Sql = Sql & "AND (Ligne_Tableau_fils.Id_IndiceProjet = Connecteurs.Id_IndiceProjet) "
    Sql = Sql & "GROUP BY Ligne_Tableau_fils.Id_IndiceProjet, Ligne_Tableau_fils.SECT, Ligne_Tableau_fils.VOI2,  "
    Sql = Sql & "Ligne_Tableau_fils.APP2, Connecteurs.CONNECTEUR "
    Sql = Sql & "Having Ligne_Tableau_fils.Id_IndiceProjet = " & Id_IndiceProjet & " And Ligne_Tableau_fils.VOI2 Is Not Null "
    Sql = Sql & "And Ligne_Tableau_fils.APP2 Is Not Null And Connecteurs.Connecteur =  '" & Connecteur & "' "
    Sql = Sql & "ORDER BY Ligne_Tableau_fils.APP2;"

DoEvents
Set Rs2 = Con.OpenRecordSet(Sql)
DoEvents
While Rs2.EOF = False
RsFils.AddNew
RsFils!Id_IndiceProjet = Rs2!Id_IndiceProjet
RsFils!SECT = ConvertTxtAsDouble("" & Rs2!SECT)
RsFils!VOI = "" & Rs2!VOIS
RsFils!App = "" & Rs2!Cod_App
RsFils!Connecteur = "" & Rs2!Connecteur
RsFils.Update
DoEvents
    Rs2.MoveNext
Wend
 DoEvents
'    Sql = "INSERT INTO TempFillesSection ( Id_IndiceProjet, VOI, SECT, APP ) IN '" & Con.RetournDbName("MDB") & "' "
'    Sql = Sql & "SELECT TempFillesSection.Id_IndiceProjet, TempFillesSection.VOI, Sum(TempFillesSection.SECT) AS SommeDeSECT, TempFillesSection.APP "
'    Sql = Sql & "FROM TempFillesSection "
'    Sql = Sql & "GROUP BY TempFillesSection.Id_IndiceProjet, TempFillesSection.VOI, TempFillesSection.APP "
'    Sql = Sql & "HAVING TempFillesSection.Id_IndiceProjet=" & Id_IndiceProjet & ";"
'    DoEvents
'Con.Execute Sql
' MySeconde 2
'
'      sql = "SELECT Rq_Alveole.* "
'    sql = sql & "FROM Rq_Alveole "
'    sql = sql & "WHERE Rq_Alveole.Référence= '" & Connecteur & "';"
'    Set Rs2 = Con.OpenRecordSet(sql)
'
DoEvents
    
    
    Sql = "SELECT Rq_Alveole.Référence, Rq_Alveole.[Nb Alvé], Rq_Alveole.Voie, Rq_Alveole.Famille, Rq_Alveole.[Famille Lib],  "
    Sql = Sql & "Rq_Alveole.[Alvé Réf], Rq_Alveole.Qté, Rq_Alveole.[Prix U], Rq_Alveole.[Prix Total], Rq_Alveole.[Alvé Réf Fourr],  "
    Sql = Sql & "Rq_Alveole.[Alvéole Mini en mm2], Rq_Alveole.[Alvéole Maxi en mm2] "
    Sql = Sql & "FROM Rq_Alveole "
    Sql = Sql & "WHERE Rq_Alveole.Id_IndiceProjet=" & Id_IndiceProjet & " AND Rq_Alveole.Référence='" & Connecteur & "' "
    Sql = Sql & "GROUP BY Rq_Alveole.Référence, Rq_Alveole.[Nb Alvé], Rq_Alveole.Voie, Rq_Alveole.Famille, Rq_Alveole.[Famille Lib],  "
    Sql = Sql & "Rq_Alveole.[Alvé Réf], Rq_Alveole.Qté, Rq_Alveole.[Prix U], Rq_Alveole.[Prix Total], Rq_Alveole.[Alvé Réf Fourr],  "
    Sql = Sql & "Rq_Alveole.[Alvéole Mini en mm2], Rq_Alveole.[Alvéole Maxi en mm2], Rq_Alveole.Id_IndiceProjet, Rq_Alveole.Référence; "
    
    Sql = "SELECT T_Alve_Eboutique.[Alvé Réf],T_Alve_Eboutique.[Famille Lib], T_Alve_Eboutique.Mini, T_Alve_Eboutique.Maxi,  "
    Sql = Sql & "T_Alve_Eboutique.[Alvé Réf Fourr],0 as Qté, T_Alve_Eboutique.[Prix u] "
    Sql = Sql & "FROM TempFillesSection, T_LientConnecteur INNER JOIN T_Alve_Eboutique ON T_LientConnecteur.Refclip = T_Alve_Eboutique.[Alvé Réf] "
    Sql = Sql & "GROUP BY T_Alve_Eboutique.[Alvé Réf], T_Alve_Eboutique.Mini, T_Alve_Eboutique.Maxi,  "
    Sql = Sql & "T_Alve_Eboutique.[Famille Lib], T_Alve_Eboutique.[Alvé Réf Fourr], T_Alve_Eboutique.[Prix u],  "
    Sql = Sql & "T_LientConnecteur.RefConnecteur, T_Alve_Eboutique.Mini, T_Alve_Eboutique.Maxi, TempFillesSection.APP, TempFillesSection.SECT "
    Sql = Sql & "HAVING T_LientConnecteur.RefConnecteur='" & Connecteur & "' AND T_Alve_Eboutique.Mini<=[SECT]  "
    Sql = Sql & "AND T_Alve_Eboutique.Maxi>=[SECT]  "
    Sql = Sql & "AND TempFillesSection.Id_IndiceProjet=" & Id_IndiceProjet & ";"

    Sql = "SELECT distinct  T_Alve_Eboutique.[Alvé Réf], T_Alve_Eboutique.[Famille Lib], T_Alve_Eboutique.[Alvé Réf Fourr],  "
    Sql = Sql & "T_Alve_Eboutique.Mini as [Alvéole Mini en mm2], T_Alve_Eboutique.Maxi as [Alvéole Maxi en mm2], 0 AS Qté, 0 AS [Alvé Prix u], 0 AS [Prix Total] "
    Sql = Sql & "FROM TempFillesSection, T_LientConnecteur INNER JOIN T_Alve_Eboutique  "
    Sql = Sql & "ON T_LientConnecteur.Refclip = T_Alve_Eboutique.[Alvé Réf] "
    Sql = Sql & "Where T_Alve_Eboutique.Mini <= [SECT] "
    Sql = Sql & "And T_Alve_Eboutique.Maxi >= [SECT] "
    Sql = Sql & "And T_LientConnecteur.RefConnecteur = '" & Connecteur & "' "
    Sql = Sql & "And TempFillesSection.Id_IndiceProjet = " & Id_IndiceProjet & " "
    Sql = Sql & "GROUP BY TempFillesSection.APP, TempFillesSection.VOI,T_Alve_Eboutique.[Alvé Réf], T_Alve_Eboutique.[Famille Lib],  "
    Sql = Sql & "T_Alve_Eboutique.[Alvé Réf Fourr], T_Alve_Eboutique.[Alvé Fornisseur],  "
    Sql = Sql & "T_Alve_Eboutique.Mini, T_Alve_Eboutique.Maxi, 0, 0, 0, TempFillesSection.SECT,  "
    Sql = Sql & "T_LientConnecteur.RefConnecteur, TempFillesSection.Id_IndiceProjet;"

'Modifier par RD
 Sql = "SELECT DISTINCT T_Lien_Connecteur_Clip.Refclip AS [Alvé Réf],lst21.CatName AS [Famille Lib], con_contacts.txt3 AS [Alvé Réf Fourr],  "
    Sql = Sql & " lst22.CatName AS [Alvéole Mini en mm2], lst23.CatName AS  "
    Sql = Sql & "[Alvéole Maxi en mm2], 0 AS Qté, 0 AS [Prix U], 0 AS [Prix Total] "
    Sql = Sql & "FROM ((((TempFillesSection INNER JOIN T_Lien_Connecteur_Clip ON TempFillesSection.CONNECTEUR =  "
    Sql = Sql & "T_Lien_Connecteur_Clip.RefConnecteur) INNER JOIN con_contacts ON T_Lien_Connecteur_Clip.Refclip =  "
    Sql = Sql & "con_contacts.txt1) LEFT JOIN lst21 ON con_contacts.lst21 = lst21.CatID) INNER JOIN lst22  "
    Sql = Sql & "ON con_contacts.lst22 = lst22.CatID) INNER JOIN lst23 ON con_contacts.lst23 = lst23.CatID "
    Sql = Sql & "WHERE T_Lien_Connecteur_Clip.RefConnecteur= '" & Connecteur & "' "
    Sql = Sql & "AND TempFillesSection.Id_IndiceProjet= " & Id_IndiceProjet & " "
    Sql = Sql & "GROUP BY T_Lien_Connecteur_Clip.Refclip,  lst21.CatName,con_contacts.txt3, lst22.CatName,  "
    Sql = Sql & "lst23.CatName, 0, 0, 0, TempFillesSection.APP, TempFillesSection.VOI,  "
    Sql = Sql & "T_Lien_Connecteur_Clip.RefConnecteur, TempFillesSection.Id_IndiceProjet "
    Sql = Sql & "HAVING (((lst22.CatName)<=Sum([SECT])) AND ((lst23.CatName)>=Sum([SECT])));"

    
    DoEvents
 Set Rs2 = Con.OpenRecordSet(Sql)
' ChargeLienObjet "RefAlve", DefinirChemienComplet(TableauPath.Item("PathServer"), TableauPath("Eb_CONNECTIQUE")), NumFieldsAlveole, NumFieldsJoint, "" & Connecteur, MySeet, CLng(Row)
 DoEvents
'    MyWorkbook.Application'.Visible = True
  
   Set TableauAlve = Nothing
   
 DoEvents
   If Rs2.EOF = False Then
        AlveolOk = True
        TableauAlve = Rs2.GetRows
        DoEvents
        ReDim tableauAlve2(UBound(TableauAlve), 1)
        DoEvents
        I = 0
        For I = 0 To UBound(TableauAlve, 2)
            On Error Resume Next
            a = ""
            a = TbAlve(TableauAlve(3, I))
            If Err <> 0 Then
                Err.Clear
                TbAlve.Add I, TableauAlve(0, I)
            End If
            tableauAlve2(TbAlve(TableauAlve(3, I)), 0) = "" & TableauAlve(1, I) & ": "
             tableauAlve2(TbAlve(TableauAlve(3, I)), 1) = tableauAlve2(TbAlve(TableauAlve(3, I)), 1) & "" & TableauAlve(2, I) & "(_____), "
             DoEvents
        Next
        Txt = ""
        For I = LBound(tableauAlve2) To UBound(tableauAlve2)
            If Trim("" & tableauAlve2(I, 1)) <> "" Then
           Txt = Txt & tableauAlve2(I, 0) & tableauAlve2(I, 1) & ";"
           
            Debug.Print Txt
            
            End If
            DoEvents
        Next
    Else
        Txt = ""
        ReDim tableauAlve2(0, 1)
     End If

Txt = Replace(Txt, ",;", "; ")
Txt = Replace(Txt, ", ;", "; ")
Debug.Print Txt
'txt = Replace(txt, ":", "")
''MyEtiquette.RenseigneChamp "Famille", "" & Txt
' MySeet.Application'.Visible = True
'  ReDim tableauAlve(1, 1) Famille
  Dim T_Alve() As String
' ReDim T_Alve(Bound(tableauAlve), 1)
  Rs2.Requery
  NumFieldsAlveole = NumFieldsJoint + Rs2.Fields.Count - 1
    For I = 0 To Rs2.Fields.Count - 1
        If iTotal < I + NumFieldsJoint Then iTotal = I + NumFieldsJoint
          MySeet.Cells(5, I + NumFieldsJoint) = Replace(Replace(Replace(Rs2(I).Name, "Qté", "Alvé Qté"), "Prix U", "Alvé Prix U"), "Prix Total", "Alvé Prix Total")
           If Rs2(I).Name = "Qté" Or Rs2(I).Name = "Prix U" Then
           MySeet.Cells(Row, I + NumFieldsJoint) = 0
         End If
         If Rs2(I).Name = "Prix Total" Then
            MySeet.Cells(Row, I + NumFieldsJoint) = "=(RC[-1]*RC[-2])"
         End If
         
    Next
     While Rs2.EOF = False
   
            For I = 0 To Rs2.Fields.Count - 1
            
                If Rs2(I).Name = "Qté" Or Rs2(I).Name = "Prix U" Or Rs2(I).Name = "Prix Total" Or Rs2(I).Name = "Prix u" Or Rs2(I).Name = "Alvé Prix u" Then
                Else
                    If UCase(Rs2(I).Name) = UCase("Voie") Then
                        MySeet.Cells(Row, I + NumFieldsJoint) = MySeet.Cells(Row, I + NumFieldsJoint) + Val(Replace("" & Rs2(I - 1), Chr(13), ""))
                    Else
                    MySeet.Cells(Row, I + NumFieldsJoint) = MySeet.Cells(Row, I + NumFieldsJoint) & Chr(10) & Replace("" & Rs2(I), Chr(13), "")
                    End If
                End If
        Next
           
             Rs2.MoveNext
        Wend
End If
      DoEvents
''      MyWord.Application'.Visible = True
'      For I = 'MyEtiquette.TableMin To 'MyEtiquette.TableMax
'      a = 'MyEtiquette.RetournEtiquette(I)
'        CreerEtiquette 'MyEtiquette.RetournEtiquette(I)
'      Next
'      MyWord'.Visible = True
'      CreerEtiquette
    Rs.MoveNext
Wend

Set Rs2 = Con.CloseRecordSet(Rs2)
'Con.CloseConnection
Set Rs = Con.CloseRecordSet(Rs)

Sql = "SELECT  T_indiceProjet.Ensemble, T_indiceProjet.Equipement,[Li] & '_' & [LI_Indice] AS Liste, "
Sql = Sql & "[PI] & '_' & [PI_Indice] AS Piece, T_indiceProjet.Client, T_indiceProjet.CleAc "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Id_IndiceProjet & ";"
Dim RsEntetePage As Recordset
Set RsEntetePage = Con.OpenRecordSet(Sql)
Set MyRange = MySeet.Range("A5").CurrentRegion
'MyWorkbook.Application'.Visible = True
MySeet.Range("A2") = "TOTAL"
MySeet.Range("D2") = "SOUS TOTAL"
MySeet.Range("N2") = "SOUS TOTAL"
MySeet.Range("W2") = "SOUS TOTAL"
MySeet.Range("AH2") = "SOUS TOTAL"

FormatExcelPlage MySeet.Range("A2"), 15, False, True, xlCenter, xlCenter
FormatExcelPlage MySeet.Range("D2"), 15, False, True, xlCenter, xlCenter
FormatExcelPlage MySeet.Range("N2"), 15, False, True, xlCenter, xlCenter
FormatExcelPlage MySeet.Range("W2"), 15, False, True, xlCenter, xlCenter
FormatExcelPlage MySeet.Range("AH2"), 15, False, True, xlCenter, xlCenter

R1 = MySeet.Range(MyRange(2, MyRange.Columns.Count).Address).Row
R2 = MySeet.Range(MyRange(MyRange.Rows.Count, MyRange.Columns.Count).Address).Row
MySeet.Range("E2").FormulaR1C1 = "=SUBTOTAL(9,R[" & R1 - 2 & "]C:R[" & R2 - 2 & "]C)"
MySeet.Range("O2").FormulaR1C1 = "=SUBTOTAL(9,R[" & R1 - 2 & "]C:R[" & R2 - 2 & "]C)"
MySeet.Range("X2").FormulaR1C1 = "=SUBTOTAL(9,R[" & R1 - 2 & "]C:R[" & R2 - 2 & "]C)"
MySeet.Range("AI2").FormulaR1C1 = "=SUBTOTAL(9,R[" & R1 - 2 & "]C:R[" & R2 - 2 & "]C)"
MySeet.Range("B2").FormulaR1C1 = "=SUM(RC[1]:RC[" & iTotal & "])"

FormatExcelPlage MySeet.Range("E2"), 2, False, True, xlCenter, xlCenter
FormatExcelPlage MySeet.Range("O2"), 2, False, True, xlCenter, xlCenter
FormatExcelPlage MySeet.Range("X2"), 2, False, True, xlCenter, xlCenter
FormatExcelPlage MySeet.Range("AI2"), 2, False, True, xlCenter, xlCenter
FormatExcelPlage MySeet.Range("B2"), 2, False, True, xlCenter, xlCenter


MiseEnPage MySeet, MySeet.Range("A5").CurrentRegion, "Affaire : " & RsEntetePage!CleAc & Chr(10) & _
 _
     "Pièce : " & RsEntetePage!Piece & Chr(10) & "Liste : " & RsEntetePage!Liste, vbCrLf & "Câblage : " & Replace("" & RsEntetePage!Ensemble, vbCrLf, " ") & Chr(10) & "Equipement : " & Replace("" & RsEntetePage!Equipement, vbCrLf, " ") & Chr(10) & _
     "" _
    , "ENCELADE" & Chr(10) & "Client : " & RsEntetePage!Client & Chr(10) & Format(Date, "dd-mmm-yyyy"), "", "&P/&N", "", 51, "B6", True, 2, True

    MaJEncadreXls MySeet.Range("A5").CurrentRegion, xlThin, xlThin, xlHairline, xlHairline

Sql = "SELECT RqCartouche.* "
Sql = Sql & "FROM RqCartouche "
 Sql = Sql & "WHERE T_indiceProjet.Id=" & Id_IndiceProjet & ";"
Set Rs = Con.OpenRecordSet(Sql)

'NomenclatureHabillage Id_IndiceProjet, MyWorkbook
 AfficheErreur PathPl, EnteteCartouche("" & Rs.Fields("Ensemble"), "" & Rs.Fields("PL_Indice"), "" & Rs.Fields("Li"))
 
'MyWordSaveAs PathPl, Save
 If IdFils <> 0 Then
        Sql = "SELECT RqCartouche.* "
        Sql = Sql & "FROM RqCartouche "
        Sql = Sql & "WHERE T_indiceProjet.Id=" & IdFils & ";"
        Set Rs2 = Con.OpenRecordSet(Sql)
         PathPl2 = PathArchive(TableauPath("PathArchiveAutocad"), "" & Rs2!Client, "" & Rs2!CleAc, "" & Rs2!Pieces, "Li", Rs2.Fields("Li"), IdFils, Rs2.Fields("PI_Indice"), Rs2.Fields("LI_Indice"), Rs2!Version)
       Racourci "" & PathPl2 & "_ETIQUETTE", PathPl & "_ETIQUETTE", "DOC"
       Racourci "" & PathPl2 & "_ETIQUETTE_MARQUAGE", PathPl & "_ETIQUETTE_MARQUAGE", "DOC"
    End If
Nomenclature3 = True
insertExelAccess MySeet, "T_Nomenclature", 5, Id_IndiceProjet
End Function

'Sub insertExelAccess(MySheet As EXCEL.Worksheet, Table As String, RowStart As Long, Id_IndiceProjet As Long)
'Dim Sql As String
'Dim SqlValue As String
'Dim Myrange As Range
'Dim Rs As Recordset
'On Error GoTo 0
'Sql = "DELETE " & Table & ".* FROM " & Table & " WHERE " & Table & ".Id_IndiceProjet=" & Id_IndiceProjet & ";"
'Con.Execute Sql
'Set Rs = Con.OpenRecordSet("SELECT " & Table & ".* FROM " & Table & " WHERE " & Table & ".ID=0;")
'
'Set Myrange = MySheet.Cells(RowStart, 1).CurrentRegion
''Myrange.Application'.Visible = True
'Sql = "INSERT INTO " & Table & " ( Id_IndiceProjet, "
'
'For I = 1 To Myrange.Columns.Count
'    Sql = Sql & "[" & Myrange(1, I) & "],"
'Next
'Sql = Left(Sql, Len(Sql) - 1) & ") Values (" & Id_IndiceProjet & ","
'
'For I = 2 To Myrange.Rows.Count
'    SqlValue = ""
'        For I2 = 1 To Myrange.Columns.Count
''        Debug.Print Myrange(I, I2).Address
''       Debug.Print Myrange(1, I2).Value & " = " & MySheet.Range(Myrange(I, I2).Address).FormulaR1C1
''       Myrange.Application'.Visible = True
'
'Debug.Print Myrange(1, I2).Value & " : " & Myrange(1, I2).Value; a; " " & "" & Myrange(I, I2).FormulaR1C1
'
'        Select Case Rs(Myrange(1, I2).Value).Type
'
'        Case 202
'            SqlValue = SqlValue & "'" & MyReplace("" & Myrange(I, I2).FormulaR1C1) & "',"
'        Case 203
'            SqlValue = SqlValue & "'" & MyReplace("" & Myrange(I, I2).FormulaR1C1) & "',"
'        Case 5
'            SqlValue = SqlValue & Replace(Val(Replace("" & MyReplace(Myrange(I, I2).FormulaR1C1), ",", ".")), ",", ".") & ","
'        Case 3
'            SqlValue = SqlValue & Replace(Val(Replace("" & MyReplace(Myrange(I, I2).FormulaR1C1), ",", ".")), ",", ".") & ","
'        Case Else
'            MsgBox ""
'        End Select
'    Next
'    SqlValue = Left(SqlValue, Len(SqlValue) - 1) & ");"
'    Con.Execute Sql & SqlValue
'
'
'Next
'
'Set Rs = Con.CloseRecordSet(Rs)
'End Sub


Sub NomenclatureHabillage(Id_IndiceProjet As Long, MyWorkbook As EXCEL.Workbook)
Dim Sql As String
Dim Rs As Recordset
Dim MySheet As EXCEL.Worksheet
Dim MyRange As Range
Dim L As Long
'MyWorkbook.Application'.Visible = True
Sql = "SELECT T_Noeuds.DESIGN_HAB, T_Noeuds.CODE_ENC,0.0  as [PRIX U/M], Sum(T_Noeuds.LONGUEUR) AS [LONGUEUR TOTAL],'=(RC[-1]*RC[-2]) * 0.001' as [Prix Total], "
Sql = Sql & " T_Noeuds.DIAMETRE, T_Noeuds.CLASSE_T FROM T_Noeuds "
Sql = Sql & "where T_Noeuds.CODE_ENC<>'_NU'  "
Sql = Sql & "AND T_Noeuds.ACTIVER=True  "
Sql = Sql & "AND T_Noeuds.Id_IndiceProjet=" & Id_IndiceProjet & " "
Sql = Sql & "GROUP BY T_Noeuds.DESIGN_HAB, T_Noeuds.CODE_ENC, T_Noeuds.DIAMETRE,  "
Sql = Sql & "T_Noeuds.CLASSE_T, T_Noeuds.ACTIVER, T_Noeuds.Id_IndiceProjet ;"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = False Then
    Set MySheet = IsertSheet(MyWorkbook, "Appro_Habillage", True)
    Set MyRange = MySheet.Cells(1, 1).CurrentRegion
    For I = 0 To Rs.Fields.Count - 1
        MyRange(1, I + 1) = Rs.Fields(I).Name
        
    Next
    L = 1
    While Rs.EOF = False
    L = L + 1
    For I = 0 To Rs.Fields.Count - 1
     If Rs.Fields(I).Name = "Prix Total" Then
        MyRange(L, I + 1).FormulaR1C1 = "" & Rs.Fields(I).Value
     Else
        MyRange(L, I + 1) = "" & Rs.Fields(I).Value
     End If
    Next
        Rs.MoveNext
    Wend
End If

Sql = "SELECT  T_indiceProjet.Ensemble, T_indiceProjet.Equipement,[Li] & '_' & [LI_Indice] AS Liste, "
Sql = Sql & "[PI] & '_' & [PI_Indice] AS Piece, T_indiceProjet.Client, T_indiceProjet.CleAc "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Id_IndiceProjet & ";"
Sql = "SELECT  T_indiceProjet.Ensemble, T_indiceProjet.Equipement,[Li] & '_' & [LI_Indice] AS Liste, "
Sql = Sql & "[PI] & '_' & [PI_Indice] AS Piece, T_indiceProjet.Client, T_indiceProjet.CleAc "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Id_IndiceProjet & ";"

Set Rs = Con.OpenRecordSet(Sql)
If PortraitPaysage = 0 Then PortraitPaysage = 2
MiseEnPage MySheet, MySheet.Range("A1").CurrentRegion, "Affaire : " & Rs!CleAc & Chr(10) & _
    "Câblage : " & Replace("" & Rs!Ensemble, vbCrLf, " ") & Chr(10) & _
     "Pièce : " & Rs!Piece, vbCrLf & "Equipement : " & Replace("" & Rs!Equipement, vbCrLf, " ") & Chr(10) & _
     "Liste : " & Rs!Liste _
    , "ENCELADE" & Chr(10) & "Client : " & Rs!Client & Chr(10) & Format(Date, "dd-mmm-yyyy"), "", "&P/&N", "", 51, "A2", True, PortraitPaysage, True

  MaJEncadreXls MySheet.Range("A1").CurrentRegion, xlThin, xlThin, xlHairline, xlHairline
  
Set Rs = Con.CloseRecordSet(Rs)

End Sub

Function EnteteCartouche(varProjet As String, varIndice As String, Plan As String)
    Dim Txt
    Dim txt2
    Dim Mysapce
   
    
     Mysapce = Space(78)
      
          Txt = "             ******************************************************************" & vbCrLf
    Txt = Txt & "             * Listes des erreurs survenues lors de l'exécution de la macro : *" & vbCrLf
    Txt = Txt & "             * Créer une Liste                                                *" & vbCrLf
         txt2 = "             * Projet : " & Replace(varProjet, vbCrLf, " ")
    Txt = Txt & txt2 & Left(Mysapce, Len(Mysapce) - Len(txt2)) & "*" & vbCrLf
         txt2 = "             * LI : " & Plan & " Indice : " & varIndice
    Txt = Txt & txt2 & Left(Mysapce, Len(Mysapce) - Len(txt2)) & "*" & vbCrLf
         txt2 = "             *"
    Txt = Txt & txt2 & Left(Mysapce, Len(Mysapce) - Len(txt2)) & "*" & vbCrLf
         txt2 = "             * Nombre d'erreur(s) : " & NbError
    Txt = Txt & txt2 & Left(Mysapce, Len(Mysapce) - Len(txt2)) & "*" & vbCrLf
    Txt = Txt & "             ******************************************************************" & vbCrLf
    Txt = Txt & vbCrLf
    EnteteCartouche = Txt
End Function

Sub ChargeLienObjet(MyRef As String, Db As String, NumEnCour As Long, NumMoinsUn As Long, Connecteur As String, MySeet As Worksheet, Row As Long, Optional RsSource As Recordset, Optional Db2 As String)
Dim RsWhere As Recordset
Dim TxtWhere As String
Dim Sql As String
Dim Rs2 As Recordset
Dim SplitConnecteur
Sql = "SELECT T_LientConnecteur.RefConnecteur, T_LientConnecteur." & MyRef & " " 'RefBouchon "
    Sql = Sql & "FROM T_LientConnecteur "
    Sql = Sql & "WHERE T_LientConnecteur.RefConnecteur='" & Connecteur & "' "
    Sql = Sql & "AND T_LientConnecteur." & MyRef & " Is Not Null "
    Sql = Sql & " GROUP BY T_LientConnecteur.RefConnecteur,  T_LientConnecteur." & MyRef & ";"
        
        Set RsWhere = Con.OpenRecordSet(Sql)
        TxtWhere = ""
        If RsWhere.EOF = False Then
            While RsWhere.EOF = False
                TxtWhere = TxtWhere & "con_contacts.txt1='" & RsWhere(MyRef) & "' OR "
                RsWhere.MoveNext
            Wend
        End If
        TxtWhere = Trim("" & TxtWhere)
         If TxtWhere <> "" Then
            TxtWhere = Left(TxtWhere, Len(TxtWhere) - 2)
             End If
            TxtWhere = " " & TxtWhere & " "
'            'Con.OpenConnetion Db '"\\10.30.0.5\production\Cablage-production\AutoCable\Access\eBoutique\Encelade_BOUCHONS.mdb")
            SplitConnecteur = Split(Connecteur & "§", "§")
    Select Case MyRef
                Case "RefConnecteur"
                Sql = "SELECT val('' & TXT9)  as [Prix U], 0 as [Prix Total],con_contacts.TXT59  as [Nb Voies] "
                Sql = Sql & "From con_contacts IN '" & Db & "' "
                Sql = Sql & "WHERE con_contacts.txt1='" & Connecteur & "';"
                Case "RefBouchon"
                    Sql = "SELECT con_contacts.txt1 as [Ref Bouch], 0 as Qté ,  Val('' & [txt9]) AS [Prix U], 0 as [Prix Total], con_contacts.mem1 AS [Lib Bouch],  "
                    Sql = Sql & "lst9.CatName AS [Bouch Fourr], con_contacts.txt3 AS [Bouch Réf Four] "
                    Sql = Sql & "FROM (con_contacts INNER JOIN lst9 ON con_contacts.lst9 = lst9.CatID)  "
                    Sql = Sql & "LEFT JOIN lst6 ON con_contacts.lst6 = lst6.CatID IN'" & Db & "' "
                    Sql = Sql & "WHERE lst9.CatName<>'(Sélectionner)' "
                    If Trim(TxtWhere) <> "" Then
                       Sql = Sql & "AND (" & TxtWhere & ")"
                    End If
                    Sql = Sql & ";"
            Case "RefCapot"
                Sql = "SELECT con_contacts.txt1 as [Ref Capot] "
                Sql = Sql & "FROM con_contacts  IN '" & Db & "' "
                If Trim(TxtWhere) <> "" Then
                    Sql = Sql & "WHERE  (" & TxtWhere & ")"
                End If
                Sql = Sql & ";"
'
          Case "RefVerrou"
                Sql = "SELECT con_contacts.txt1 as [Ref Verrou] "
                Sql = Sql & "FROM con_contacts IN '" & Db & "' "
                If Trim(TxtWhere) <> "" Then
                    Sql = Sql & "WHERE  (" & TxtWhere & ")"
                End If
                Sql = Sql & ";"
         Case "RefJoint"
                Sql = "SELECT con_contacts.txt1 as [Ref Joint],0 as Qté, Val('' & [txt9]) AS [Prix U],0 as [Prix Total],con_contacts.mem1 AS [Lib Joint], "
                 Sql = Sql & "lst9.CatName AS [Joint Four], con_contacts.txt3 AS [Joint Four Réf] "
                 Sql = Sql & "FROM (con_contacts INNER JOIN lst9 ON con_contacts.lst9 = lst9.CatID)  IN '" & Db & "' "
                  Sql = Sql & "WHERE lst9.CatName<>'(Sélectionner)'  "
                If Trim(TxtWhere) <> "" Then
                    Sql = Sql & "AND  (" & TxtWhere & ")"
                End If
                Sql = Sql & ";"
    End Select
   
     Set Rs2 = Con.OpenRecordSet(Sql)
   

'    NumFieldsBouchon = NumFieldsFournisseur + Rs2.Fields.Count
    NumEnCour = NumMoinsUn + Rs2.Fields.Count
    For I = 0 To Rs2.Fields.Count - 1
    
          MySeet.Cells(5, I + NumMoinsUn) = Replace(Replace(Replace(Replace(Replace(Rs2(I).Name, "Qté", Replace(MyRef, "Ref", "") & " Qté"), "Prix U", Replace(MyRef, "Ref", "") & " Prix U"), "Prix Total", Replace(MyRef, "Ref", "") & " Prix Total"), "Connecteur Prix Total", "Prix Total"), "Connecteur Prix U", "Prix U")
         
         If InStr(1, UCase(Rs2(I).Name), UCase("Qté")) <> 0 Or InStr(1, UCase(Rs2(I).Name), UCase("Prix U")) <> 0 Then
          
            MySeet.Cells(Row, I + NumMoinsUn) = 0
         End If
         If Rs2(I).Name = "Prix Total" Then
            MySeet.Cells(Row, I + NumMoinsUn).FormulaR1C1 = "=(RC[-1]*RC[-2])"
         End If
    Next
    
    If Rs2.EOF = True And MyRef = "RefConnecteur" Then
        'Con.OpenConnetion Db2
         Set Rs2 = Con.OpenRecordSet(Sql)
         TxtWhere = "?"
      End If
      If Trim("" & TxtWhere) <> "" Then
     While Rs2.EOF = False
   
            For I = 0 To Rs2.Fields.Count - 1
                  
                   
                             If Rs2(I).Name = "Qté" Or Rs2(I).Name = "Prix U" Or Rs2(I).Name = "Prix Total" Then
                                If (Rs2(I).Name = "Prix U" And MyRef = "RefConnecteur") Then
                                    MySeet.Cells(Row, I + NumMoinsUn) = Rs2(I)
                                End If
                                
                              Else
                              If (Rs2(I).Name = "PU HT") Then
                                 MySeet.Cells(Row, I + NumMoinsUn) = Rs2(I)
                              Else
                                MySeet.Cells(Row, I + NumMoinsUn) = MySeet.Cells(Row, I + NumMoinsUn) & Chr(10) & Replace(Replace("" & Rs2(I), Chr(13), ""), Chr(10), " ")
                              End If
                  End If
                  Next
             Rs2.MoveNext
        Wend
      End If
       Set Rs2 = Con.CloseRecordSet(Rs2)
'       'Con.CloseConnection
End Sub
