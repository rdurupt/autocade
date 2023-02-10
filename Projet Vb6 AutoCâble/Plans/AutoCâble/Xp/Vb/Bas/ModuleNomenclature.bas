Attribute VB_Name = "ModuleNomenclature"


Public Function Nomenclature(Id_IndiceProjet As Long, Optional PathPl As String) As Boolean
Dim ConCommposants As New Ado
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
Dim Myrange As Range
Dim PathModelWord As String
Dim PIE(1) As String
Dim Ensemble(1) As String
Dim i As Long
Dim iTotal As Long
Dim MySeet As Worksheet
Set MySeet = IsertSheet(MyWorkbook, "Appro Connectique", True)
DbOk = ConCommposants.OpenConnetion(DbCatalogue) '"U:\Librairies\Plans\AutoCâble\Xp\Access\Catalogue Renault.mdb"
DeleteRow MySeet, True

Sql = "SELECT [PI] & '_' & [PI_Indice] AS PIE, T_indiceProjet.Ensemble "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Id_IndiceProjet & ";"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = False Then
    PIE(0) = "PI": PIE(1) = "" & Rs!PIE
    Ensemble(0) = "Ensemble": Ensemble(1) = "" & Rs!Ensemble
End If

Sql = "SELECT Rq_Compte_Connecteur_IdPices.CONNECTEUR, Rq_Compte_Connecteur_IdPices.[Qté],0 as [Prix U],'=(RC[-1]*RC[-2]) ' as [Prix Total] "
Sql = Sql & "FROM Rq_Compte_Connecteur_IdPices "
Sql = Sql & "WHERE Rq_Compte_Connecteur_IdPices.CONNECTEUR<>'NEANT' "
Sql = Sql & "AND Rq_Compte_Connecteur_IdPices.Id_IndiceProjet=" & Id_IndiceProjet & " " '94 "
Sql = Sql & "ORDER BY Rq_Compte_Connecteur_IdPices.CONNECTEUR;"
Set Rs = Con.OpenRecordSet(Sql)
NumFieldsConnecteur = Rs.Fields.Count
    If Rs.EOF = True Then Exit Function
    
Set MyWord = New Word.Application

'MyWord.Visible = True

PathModelWord = TableauPath.Item("PathModelWord")
         If Left(PathModelWord, 2) <> "\\" And Left(PathModelWord, 1) = "\" Then PathModelWord = TableauPath.Item("PathServer") & PathModelWord
          If Right(PathModelWord, 2) = "\\" Then PathModelWord = Mid(PathModelWord, 1, Len(PathModelWord) - 1)
Set MyWordDoc = WordNewDoc(PathModelWord)
          
'MyEcel.Visible = True
    
While Rs.EOF = False
NbLigne = NbLigne + 1
Rs.MoveNext
Wend
Rs.MoveFirst
'MyWorkbook.Application.Visible = True
For i = 0 To Rs.Fields.Count - 1
    MySeet.Cells(5, i + 1) = Rs(i).Name
Next
Row = 5
 FormBarGrah.ProgressBar1.Value = 0
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
      MySeet.Cells(Row, 4).FormulaR1C1 = "" & Rs(3)
      Sql = "SELECT Connecteurs.CODE_APP, Connecteurs.DESIGNATION "
    Sql = Sql & "FROM Connecteurs "
    Sql = Sql & "WHERE Connecteurs.Id_IndiceProjet=" & Id_IndiceProjet & " "
    Sql = Sql & "AND Connecteurs.CONNECTEUR='" & Connecteur & "' "
    Sql = Sql & "AND Connecteurs.[O/N]=False;"
     Set Rs2 = Con.OpenRecordSet(Sql)
        NumFieldsConnecteurApp = NumFieldsConnecteur + Rs2.Fields.Count - 1
     MyEtiquette.PrpareEtiqet Rs2.GetString
     MyEtiquette.RenseigneChamp "Connecteur", "" & Connecteur
     MyEtiquette.RenseigneChamp "" & PIE(0), "" & PIE(1)
     MyEtiquette.RenseigneChamp "" & Ensemble(0), "" & Ensemble(1)
    For i = 1 To Rs2.Fields.Count
        MySeet.Cells(5, i + NumFieldsConnecteur) = Rs2(i - 1).Name
    
    Next
    Rs2.Requery
    If Rs2.EOF = True Then
  
        FunError 8, "" & Connecteur, "" & vbCrLf & Err.Description
    End If
    
    While Rs2.EOF = False
    
            For i = 1 To Rs2.Fields.Count
        MySeet.Cells(Row, i + NumFieldsConnecteur) = MySeet.Cells(Row, i + NumFieldsConnecteur) & Chr(10) & Replace("" & Rs2(i - 1), Chr(13), "")
         
        Next
             Rs2.MoveNext
        Wend
   If DbOk = True Then
    Sql = "SELECT Rq_Fournisseur.* "
    Sql = Sql & "FROM Rq_Fournisseur "
    Sql = Sql & "WHERE Rq_Fournisseur.[Ref Connecteur]= '" & Connecteur & "';"
    Set Rs2 = ConCommposants.OpenRecordSet(Sql)
    NumFieldsFournisseur = NumFieldsConnecteurApp + Rs2.Fields.Count - 1
    For i = 2 To Rs2.Fields.Count - 1
        MySeet.Cells(5, i + NumFieldsConnecteurApp) = Rs2(i).Name
    
    Next
    If Rs2.EOF = True Then
  
        FunError 8, "" & Connecteur, "" & vbCrLf & Err.Description
    End If
    While Rs2.EOF = False
    
            For i = 2 To Rs2.Fields.Count - 1
        MySeet.Cells(Row, i + NumFieldsConnecteurApp) = MySeet.Cells(Row, i + NumFieldsConnecteurApp) & Chr(10) & Replace("" & Rs2(i), Chr(13), "")
         
        Next
             Rs2.MoveNext
        Wend
        
        
        Sql = "SELECT Rq_Bouchon.* "
    Sql = Sql & "FROM Rq_Bouchon "
    Sql = Sql & "WHERE Rq_Bouchon.Référence= '" & Connecteur & "';"
    Set Rs2 = ConCommposants.OpenRecordSet(Sql)
   

    NumFieldsBouchon = NumFieldsFournisseur + Rs2.Fields.Count - 1
    For i = 1 To Rs2.Fields.Count - 1
          MySeet.Cells(5, i + NumFieldsFournisseur) = Rs2(i).Name
         If Rs2(i).Name = "Qté" Or Rs2(i).Name = "Prix U" Then
            MySeet.Cells(Row, i + NumFieldsFournisseur) = 0
         End If
         If Rs2(i).Name = "Prix Total" Then
            MySeet.Cells(Row, i + NumFieldsFournisseur).FormulaR1C1 = "=(RC[-1]*RC[-2])"
         End If
    Next
     While Rs2.EOF = False
   
            For i = 1 To Rs2.Fields.Count - 1
                  
                   
                             If Rs2(i).Name = "Qté" Or Rs2(i).Name = "Prix U" Or Rs2(i).Name = "Prix Total" Then
                                
                              Else
                                MySeet.Cells(Row, i + NumFieldsFournisseur) = MySeet.Cells(Row, i + NumFieldsFournisseur) & Chr(10) & Replace("" & Rs2(i), Chr(13), "")
                  End If
                  Next
             Rs2.MoveNext
        Wend
        
           
        Sql = "SELECT Rq_Capot.* "
    Sql = Sql & "FROM Rq_Capot "
    Sql = Sql & "WHERE Rq_Capot.Référence= '" & Connecteur & "';"
    Set Rs2 = ConCommposants.OpenRecordSet(Sql)
  
    NumFieldsCapot = NumFieldsBouchon + Rs2.Fields.Count - 1
    For i = 1 To Rs2.Fields.Count - 1
          MySeet.Cells(5, i + NumFieldsBouchon) = Rs2(i).Name
         
    Next
     While Rs2.EOF = False
   
            For i = 0 To Rs2.Fields.Count - 1
          MySeet.Cells(Row, i + NumFieldsBouchon) = MySeet.Cells(Row, i + NumFieldsBouchon) & Chr(10) & Replace("" & Rs2(i), Chr(13), "")
         
        Next
             Rs2.MoveNext
        Wend
        
      Sql = "SELECT Rq_Verou.* "
    Sql = Sql & "FROM Rq_Verou "
    Sql = Sql & "WHERE Rq_Verou.Référence= '" & Connecteur & "';"
    Set Rs2 = ConCommposants.OpenRecordSet(Sql)
  
   NumFieldsVerou = NumFieldsCapot + Rs2.Fields.Count - 1
    For i = 1 To Rs2.Fields.Count - 1
          MySeet.Cells(5, i + NumFieldsCapot) = Rs2(i).Name
         
    Next
     While Rs2.EOF = False
   
            For i = 0 To Rs2.Fields.Count - 1
          MySeet.Cells(Row, i + NumFieldsCapot) = MySeet.Cells(Row, i + NumFieldsCapot) & Chr(10) & Replace("" & Rs2(i), Chr(13), "")
         
        Next
             Rs2.MoveNext
        Wend
         Sql = "SELECT Rq_Joint.* "
    Sql = Sql & "FROM Rq_Joint "
    Sql = Sql & "WHERE Rq_Joint.Référence= '" & Connecteur & "';"
    Set Rs2 = ConCommposants.OpenRecordSet(Sql)
  
  NumFieldsJoint = NumFieldsVerou + Rs2.Fields.Count - 1
    For i = 1 To Rs2.Fields.Count - 1
          MySeet.Cells(5, i + NumFieldsVerou) = Rs2(i).Name
           If Rs2(i).Name = "Qté" Or Rs2(i).Name = "Prix U" Then
            MySeet.Cells(Row, i + NumFieldsVerou) = 0
           End If
           If Rs2(i).Name = "Prix Total" Then
           MySeet.Cells(Row, i + NumFieldsVerou).FormulaR1C1 = "=(RC[-1]*RC[-2])"
           End If
         
    Next
     While Rs2.EOF = False
   
            For i = 1 To Rs2.Fields.Count - 1
          
          
                 If Rs2(i).Name = "Qté" Or Rs2(i).Name = "Prix U" Or Rs2(i).Name = "Prix Total" Then
                   
                Else
                MySeet.Cells(Row, i + NumFieldsVerou) = MySeet.Cells(Row, i + NumFieldsVerou) & Chr(10) & Replace("" & Rs2(i), Chr(13), "")
                End If
            
          MyEtiquette.RenseigneChamp "" & Rs2(i).Name, "" & Rs2(i).Value
        Next
             Rs2.MoveNext
'        End If
Wend

      Sql = "SELECT Rq_Alveole.* "
    Sql = Sql & "FROM Rq_Alveole "
    Sql = Sql & "WHERE Rq_Alveole.Référence= '" & Connecteur & "';"
    Set Rs2 = ConCommposants.OpenRecordSet(Sql)
'    MyWorkbook.Application.Visible = True
  Dim TbAlve As New Collection

   Set TableauAlve = Nothing
 
   If Rs2.EOF = False Then
   TableauAlve = Rs2.GetRows
    ReDim tableauAlve2(UBound(TableauAlve), 1)
    i = 0
For i = 0 To UBound(TableauAlve, 2)
On Error Resume Next
    a = ""
    a = TbAlve(TableauAlve(3, i))
    If Err <> 0 Then
        Err.Clear
        TbAlve.Add i, TableauAlve(3, i)
    End If
    tableauAlve2(TbAlve(TableauAlve(3, i)), 0) = "" & TableauAlve(3, i) & ": "
     tableauAlve2(TbAlve(TableauAlve(3, i)), 1) = tableauAlve2(TbAlve(TableauAlve(3, i)), 1) & "" & TableauAlve(5, i) & "(_____), "
Next
txt = ""
For i = LBound(tableauAlve2) To UBound(tableauAlve2)
    If Trim("" & tableauAlve2(i, 1)) <> "" Then
   txt = txt & tableauAlve2(i, 0) & tableauAlve2(i, 1) & ";"
   
    Debug.Print txt
    End If
Next

    Else
     ReDim tableauAlve2(0, 1)
     End If

txt = Replace(txt, ",;", "; ")
Debug.Print txt
'txt = Replace(txt, ":", "")
MyEtiquette.RenseigneChamp "Famille", "" & txt
 
'  ReDim tableauAlve(1, 1) Famille
  Dim T_Alve() As String
' ReDim T_Alve(Bound(tableauAlve), 1)
  Rs2.Requery
  NumFieldsAlveole = NumFieldsJoint + Rs2.Fields.Count - 1
    For i = 1 To Rs2.Fields.Count - 1
        If iTotal < i + NumFieldsJoint Then iTotal = i + NumFieldsJoint
          MySeet.Cells(5, i + NumFieldsJoint) = Rs2(i).Name
           If Rs2(i).Name = "Qté" Or Rs2(i).Name = "Prix U" Then
           MySeet.Cells(Row, i + NumFieldsJoint) = 0
         End If
         If Rs2(i).Name = "Prix Total" Then
            MySeet.Cells(Row, i + NumFieldsJoint) = "=(RC[-1]*RC[-2])"
         End If
         
    Next
     While Rs2.EOF = False
   
            For i = 1 To Rs2.Fields.Count - 1
            
                If Rs2(i).Name = "Qté" Or Rs2(i).Name = "Prix U" Or Rs2(i).Name = "Prix Total" Then
                Else
                    If UCase(Rs2(i).Name) = UCase("Voie") Then
                        MySeet.Cells(Row, i + NumFieldsJoint) = MySeet.Cells(Row, i + NumFieldsJoint) + Val(Replace("" & Rs2(i - 1), Chr(13), ""))
                    Else
                    MySeet.Cells(Row, i + NumFieldsJoint) = MySeet.Cells(Row, i + NumFieldsJoint) & Chr(10) & Replace("" & Rs2(i), Chr(13), "")
                    End If
                End If
        Next
           
             Rs2.MoveNext
        Wend
End If
      DoEvents
'      MyWord.Application.Visible = True
      For i = MyEtiquette.TableMin To MyEtiquette.TableMax
      a = MyEtiquette.RetournEtiquette(i)
        CreerEtiquette MyEtiquette.RetournEtiquette(i)
      Next
'      CreerEtiquette
    Rs.MoveNext
Wend

Set Rs2 = ConCommposants.CloseRecordSet(Rs2)
ConCommposants.CloseConnection
Set Rs = Con.CloseRecordSet(Rs)

Sql = "SELECT  T_indiceProjet.Ensemble, T_indiceProjet.Equipement,[Li] & '_' & [LI_Indice] AS Liste, "
Sql = Sql & "[PI] & '_' & [PI_Indice] AS Piece, T_indiceProjet.Client, T_indiceProjet.CleAc "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Id_IndiceProjet & ";"
Dim RsEntetePage As Recordset
Set RsEntetePage = Con.OpenRecordSet(Sql)
Set Myrange = MySeet.Range("A5").CurrentRegion
'MyWorkbook.Application.Visible = True
MySeet.Range("A1") = "TOTAL"
MySeet.Range("C2") = "SOUS TOTAL"
MySeet.Range("M2") = "SOUS TOTAL"
MySeet.Range("V2") = "SOUS TOTAL"
MySeet.Range("AG2") = "SOUS TOTAL"

FormatExcelPlage MySeet.Range("A1"), 15, False, True, xlCenter, xlCenter
FormatExcelPlage MySeet.Range("C2"), 15, False, True, xlCenter, xlCenter
FormatExcelPlage MySeet.Range("M2"), 15, False, True, xlCenter, xlCenter
FormatExcelPlage MySeet.Range("V2"), 15, False, True, xlCenter, xlCenter
FormatExcelPlage MySeet.Range("AG2"), 15, False, True, xlCenter, xlCenter

r1 = MySeet.Range(Myrange(2, Myrange.Columns.Count).Address).Row
r2 = MySeet.Range(Myrange(Myrange.Rows.Count, Myrange.Columns.Count).Address).Row
MySeet.Range("D2").FormulaR1C1 = "=SUBTOTAL(9,R[" & r1 - 2 & "]C:R[" & r2 - 2 & "]C)"
MySeet.Range("N2").FormulaR1C1 = "=SUBTOTAL(9,R[" & r1 - 2 & "]C:R[" & r2 - 2 & "]C)"
MySeet.Range("W2").FormulaR1C1 = "=SUBTOTAL(9,R[" & r1 - 2 & "]C:R[" & r2 - 2 & "]C)"
MySeet.Range("AH2").FormulaR1C1 = "=SUBTOTAL(9,R[" & r1 - 2 & "]C:R[" & r2 - 2 & "]C)"
MySeet.Range("A2").FormulaR1C1 = "=SUM(RC[1]:RC[" & iTotal & "])"

FormatExcelPlage MySeet.Range("D2"), 2, False, True, xlCenter, xlCenter
FormatExcelPlage MySeet.Range("N2"), 2, False, True, xlCenter, xlCenter
FormatExcelPlage MySeet.Range("W2"), 2, False, True, xlCenter, xlCenter
FormatExcelPlage MySeet.Range("AH2"), 2, False, True, xlCenter, xlCenter
FormatExcelPlage MySeet.Range("A2"), 2, False, True, xlCenter, xlCenter


MiseEnPage MySeet, MySeet.Range("A5").CurrentRegion, "Affaire : " & RsEntetePage!CleAc & vbCrLf & _
 _
     "Pièce : " & RsEntetePage!Piece & vbCrLf & "Liste : " & RsEntetePage!Liste, vbCrLf & "Câblage : " & Replace("" & RsEntetePage!Ensemble, vbCrLf, " ") & vbCrLf & "Equipement : " & Replace("" & RsEntetePage!Equipement, vbCrLf, " ") & vbCrLf & _
     "" _
    , "ENCELADE" & vbCrLf & "Client : " & RsEntetePage!Client & vbCrLf & Format(Date, "dd-mmm-yyyy"), "", "&P/&N", "", 51, "B6", True, 2, True

    MaJEncadreXls MySeet.Range("A5").CurrentRegion, xlThin, xlThin, xlHairline, xlHairline


Sql = "SELECT RqCartouche.* "
Sql = Sql & "FROM RqCartouche "
 Sql = Sql & "WHERE T_indiceProjet.Id=" & Id_IndiceProjet & ";"
Set Rs = Con.OpenRecordSet(Sql)

'NomenclatureHabillage Id_IndiceProjet, MyWorkbook
 AfficheErreur PathPl, EnteteCartouche("" & Rs.Fields("Ensemble"), "" & Rs.Fields("PL_Indice"), "" & Rs.Fields("Li"))
MyWordSaveAs PathPl
 If IdFils <> 0 Then
        Sql = "SELECT RqCartouche.* "
        Sql = Sql & "FROM RqCartouche "
        Sql = Sql & "WHERE T_indiceProjet.Id=" & IdFils & ";"
        Set Rs2 = Con.OpenRecordSet(Sql)
         PathPl2 = PathArchive(PathArchiveAutocad, "" & Rs2!Client, "" & Rs2!CleAc, "" & Rs2!Pieces, "Li", Rs2.Fields("Li"), IdFils, Rs2.Fields("PI_Indice"), Rs2.Fields("LI_Indice"), Rs2!Version)
       Racourci "" & PathPl2 & "_ETIQUETTE", PathPl & "_ETIQUETTE", "DOC"
    End If
Nomenclature = True
End Function
Sub NomenclatureHabillage(Id_IndiceProjet As Long, MyWorkbook As EXCEL.Workbook)
Dim Sql As String
Dim Rs As Recordset
Dim MySheet As EXCEL.Worksheet
Dim Myrange As Range
Dim L As Long
'MyWorkbook.Application.Visible = True
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
    Set Myrange = MySheet.Cells(1, 1).CurrentRegion
    For i = 0 To Rs.Fields.Count - 1
        Myrange(1, i + 1) = Rs.Fields(i).Name
        
    Next
    L = 1
    While Rs.EOF = False
    L = L + 1
    For i = 0 To Rs.Fields.Count - 1
     If Rs.Fields(i).Name = "Prix Total" Then
        Myrange(L, i + 1).FormulaR1C1 = "" & Rs.Fields(i).Value
     Else
        Myrange(L, i + 1) = "" & Rs.Fields(i).Value
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

MiseEnPage MySheet, MySheet.Range("A1").CurrentRegion, "Affaire : " & Rs!CleAc & vbCrLf & _
    "Câblage : " & Replace("" & Rs!Ensemble, vbCrLf, " ") & vbCrLf & _
     "Pièce : " & Rs!Piece, vbCrLf & "Equipement : " & Replace("" & Rs!Equipement, vbCrLf, " ") & vbCrLf & _
     "Liste : " & Rs!Liste _
    , "ENCELADE" & vbCrLf & "Client : " & Rs!Client & vbCrLf & Format(Date, "dd-mmm-yyyy"), "", "&P/&N", "", 51, "A2", True, 2, True

  MaJEncadreXls MySheet.Range("A1").CurrentRegion, xlThin, xlThin, xlHairline, xlHairline
  
Set Rs = Con.CloseRecordSet(Rs)

End Sub

Function EnteteCartouche(varProjet As String, varIndice As String, Plan As String)
    Dim txt
    Dim txt2
    Dim Mysapce
   
    
     Mysapce = Space(78)
      
          txt = "             ******************************************************************" & vbCrLf
    txt = txt & "             * Listes des erreurs survenues lors de l'exécution de la macro : *" & vbCrLf
    txt = txt & "             * Créer une Liste                                                *" & vbCrLf
         txt2 = "             * Projet : " & Replace(varProjet, vbCrLf, " ")
    txt = txt & txt2 & Left(Mysapce, Len(Mysapce) - Len(txt2)) & "*" & vbCrLf
         txt2 = "             * LI : " & Plan & " Indice : " & varIndice
    txt = txt & txt2 & Left(Mysapce, Len(Mysapce) - Len(txt2)) & "*" & vbCrLf
         txt2 = "             *"
    txt = txt & txt2 & Left(Mysapce, Len(Mysapce) - Len(txt2)) & "*" & vbCrLf
         txt2 = "             * Nombre d'erreur(s) : " & NbError
    txt = txt & txt2 & Left(Mysapce, Len(Mysapce) - Len(txt2)) & "*" & vbCrLf
    txt = txt & "             ******************************************************************" & vbCrLf
    txt = txt & vbCrLf
    EnteteCartouche = txt
End Function

