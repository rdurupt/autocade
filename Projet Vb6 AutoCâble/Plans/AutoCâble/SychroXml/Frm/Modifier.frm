VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Modifier 
   Caption         =   "Exporter  Xml & Importer HTML:"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9120
   Icon            =   "Modifier.dsx":0000
   OleObjectBlob   =   "Modifier.dsx":030A
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Modifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public IdIndiceProjet As Long
Dim Id_Pere As Long
Dim Noquite As Boolean


Private Sub CommandButton1_Click()
Dim Sql As String
CherchPices.Charge Me, "(VerifieDate= Null  and Archiver=False) OR (IdStatus=3 and Archiver=False)"
Sql = "UPDATE T_indiceProjet SET T_indiceProjet.UserName = null "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Me.Controls("txt1").Tag & " OR T_indiceProjet.Pere=" & Me.Controls("txt1").Tag & ";"
Con.Execute Sql
Unload CherchPices
End Sub

Private Sub CommandButton2_Click()
Dim SplitPathPl
Dim MyClaseur As Workbook
Dim txtSheetName As String
Dim Sql As String
Dim SqlWher As String
Dim Rs As Recordset
Dim MySheet As Worksheet
Dim MySheet2  As Worksheet
Dim MyRange As Range
Dim Myrange2 As Range
Dim MySplit
Dim VoieUne As Boolean
Dim I As Long
Dim Ofset As Long
Dim MyVal As Double
Dim NbVoiEpisur As Long
Dim IEP As Long
Dim Txt As String
Dim SershRow As Long
Dim Equipement
Dim SplitEquipement
Dim Piece As String
Dim PathPl As String
Dim MySplit2
Dim IS_Term1 As Boolean
Dim PnG As Long
Dim PnD As Long
Dim PnT As Long
Dim ValEpisure As String
Dim RsPasTrouver As Recordset
Dim Sql2 As String
Dim MaxPin

If Trim("" & Me.txt3.Tag) = "" Then
    MsgBox "Le champ Pièce est obligatoire.", vbCritical, "Auto-Câble"
    CommandButton1_Click
    Exit Sub
End If

On Error Resume Next
Set TableauPath = funPath

Sql = "SELECT T_indiceProjet.Pere FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Me.txt1.Tag & ";"
Set Rs = Con.OpenRecordSet(Sql)

IdFils = 0

If Rs!Pere > 0 Then

    Me.Tag = Rs!Pere
Else
Me.Tag = Me.txt1.Tag
End If
Sql = "SELECT RqCartouche.* "
Sql = Sql & "FROM RqCartouche "
 Sql = Sql & "WHERE T_indiceProjet.Id=" & Me.txt1.Tag & ";"
Set Rs = Con.OpenRecordSet(Sql)
Equipement = "" & Rs!Equipement
Piece = "" & Rs!Pi & "_" & Rs!PI_Indice




If OptionButton1.Value = True Then

Set MyExcel = New EXCEL.Application
MyExcel.DisplayAlerts = False
MyExcel.Visible = True
MyExcel.DisplayAlerts = False




PathPl = PathArchive(TableauPath("PathArchiveAutocad"), "" & Rs!Client, "" & Rs!CleAc, "" & Rs!Pieces, "CatiaV5", Rs.Fields("Pi"), IdIndiceProjet, Rs.Fields("PI_Indice"), Rs.Fields("OU_Indice"), Rs!Version)
SplitPathPl = Split(PathPl, "\")
Set MyClaseur = MyExcel.Workbooks.Add(TableauPath("ModelPiCatiaV5"))
MyClaseur.Sheets("Create").Range("b5") = ""
For I = 0 To UBound(SplitPathPl) - 1

MyClaseur.Sheets("Create").Range("b5") = MyClaseur.Sheets("Create").Range("b5") & "" & SplitPathPl(I) & "\"
Next
MyClaseur.Sheets("Create").Range("b5") = Left(MyClaseur.Sheets("Create").Range("b5"), Len(MyClaseur.Sheets("Create").Range("b5")) - 1)
MyClaseur.Sheets("Create").Range("b3") = SplitPathPl(UBound(SplitPathPl))

SplitEquipement = Split(Equipement & ";", ";")

Sql = "SELECT Connecteurs.CONNECTEUR, Connecteurs.CODE_APP, Connecteurs.CODE_APP AS CODE_APP2, Connecteurs.[O/N] "
Sql = Sql & "From Connecteurs  "
Sql = Sql & "Where Connecteurs.Id_IndiceProjet = " & Me.Tag & " AND Connecteurs.ACTIVER=True "
Sql = Sql & "ORDER BY Connecteurs.CONNECTEUR;"


Sql = "SELECT Connecteurs.CONNECTEUR, Connecteurs.CODE_APP, Connecteurs.CODE_APP AS CODE_APP2, Connecteurs.[O/N] "
Sql = Sql & "FROM (Connecteurs LEFT JOIN T_Critères ON (Connecteurs.Id_IndiceProjet = T_Critères.Id_IndiceProjet)  "
Sql = Sql & "AND (Connecteurs.OPTION = T_Critères.CRITERES)) LEFT JOIN T_Critères AS T_Critères_1 ON (Connecteurs.Id_IndiceProjet =  "
Sql = Sql & "T_Critères_1.Id_IndiceProjet) AND (Connecteurs.OPTION = T_Critères_1.CODE_CRITERE) "

Sql = "SELECT  Connecteurs.CONNECTEUR, Connecteurs.CODE_APP, '' as aa ,'' as aaa ,Connecteurs.[O/N] "
Sql = Sql & "FROM Connecteurs "
'Sql = Sql & "GROUP BY Connecteurs.Id_IndiceProjet, Connecteurs.CONNECTEUR, Connecteurs.CODE_APP, Connecteurs.[O/N] "
'Sql = Sql & "Having (((Connecteurs.Id_IndiceProjet) = 376)) "
'Sql = Sql & "ORDER BY Connecteurs.CONNECTEUR;"

Sql = Sql & "Where Connecteurs.Id_IndiceProjet=" & Me.Tag & " AND Connecteurs.ACTIVER=True AND  "
Sql = Sql & "(Connecteurs.OPTION Is Null or  Connecteurs.OPTION Is Null "

Sql = Sql & "OR ';' & Connecteurs.OPTION & ';'  Like '%;tous;%'  OR ';' & Connecteurs.OPTION & ';'  Like '%;ALL;%' "

For I = 0 To UBound(SplitEquipement) - 1
    If Trim("" & SplitEquipement(I)) <> "" Then
    MySplit2 = Split(SplitEquipement(I) & "_", "_")
        Sql = Sql & "OR ';' & [Connecteurs].[OPTION] & ';'  Like '%;" & MySplit2(0) & ";%' "
        Sql = Sql & "OR ';' & Connecteurs.OPTION & ';'  Like '%;" & MySplit2(0) & ";%' "
    End If
Next
' (((';' & [T_Critères_1].[CRITERES] & ';') Like '*;tous;*')) OR (((';' & [T_Critères].[CRITERES] & ';') Like '*;BVM;*'))
Sql = Sql & ") ORDER BY Connecteurs.CONNECTEUR;"





Set Rs = Con.OpenRecordSet(Sql)
Rs.Filter = "[O/N]=false"
Set MySheet = MyClaseur.Sheets("SIC-TERM")
MySheet.Select
MySheet.Range("A2").CopyFromRecordset Rs
Rs.Requery
Rs.Filter = "[O/N]=true"
Set MySheet2 = MyClaseur.Sheets("IS")
MySheet2.Select
MySheet2.Range("A2").CopyFromRecordset Rs
MySheet.Select
Set MyRange = MySheet.Range("a1").CurrentRegion
Sql = "SELECT T_Lien_Con_Famille.Connecteur, T_Lien_Con_Famille_Voies.Voie "
Sql = Sql & "FROM T_Lien_Con_Famille INNER JOIN T_Lien_Con_Famille_Voies  "
Sql = Sql & "ON T_Lien_Con_Famille.Id = T_Lien_Con_Famille_Voies.Id_T_Lien_Con_Famille "




For I = MyRange.Rows.Count To 2 Step -1
'If I = 7 Then MsgBox ""
DoEvents
 VoieUne = False
 Ofset = 0
MySplit = MyRange(I, 1).Value
MyRange(I, 2).Select
MyRange(I, 2) = Replace("" & MyRange(I, 2), ".", "*")

MyRange(I, 3) = MyRange(I, 2)
MySplit = Split(MySplit & "§", "§")
'If UBound(MySplit) > 1 Then MsgBox ""

MyRange(I, 1).Value = MySplit(0)
'If UCase(MyRange(I, 4)) = "VRAI" Or UCase(MyRange(I, 4)) = "YES" Then
'    NbVoiEpisur = Val(Trim(Replace(MyRange(I, 1), "EP", "")))
'    NbVoiEpisur = NbVoiEpisur / 2
'    MyRange(I, 1) = "'IS-Term1"
'    For IEP = 0 To 49
'    If VoieUne = False Then
'        VoieUne = True
'        MyRange(I + Ofset, 4) = ""
'         MyRange(I + Ofset, 5) = "P" & IEP
'         MyRange(I + Ofset, 6) = MyRange(I, 2) & "." & MyRange(I + Ofset, 5)
'         Ofset = Ofset + 1
''        InsertRow Myrange, I + Ofset
''        Myrange(I + Ofset, 5) = "D" & IEP
''        Myrange(I + Ofset, 6) = Myrange(I, 2) & "." & Myrange(I + Ofset, 5)
'    Else
'
'
'        InsertRow MyRange, I + Ofset
'         MyRange(I + Ofset, 4) = ""
'         MyRange(I + Ofset, 5) = "P" & IEP
'         MyRange(I + Ofset, 6) = MyRange(I, 2) & "." & MyRange(I + Ofset, 5)
'         Ofset = Ofset + 1
''        InsertRow Myrange, I + Ofset
''        Myrange(I + Ofset, 5) = "D" & IEP
''        Myrange(I + Ofset, 6) = Myrange(I, 2) & "." & Myrange(I + Ofset, 5)
'    End If
'    Next
'Else
SqlWher = "WHERE T_Lien_Con_Famille.Connecteur='" & MySplit(0) & "';"


Sql = "SELECT Ligne_Tableau_fils.APP, Ligne_Tableau_fils.VOI, Ligne_Tableau_fils.APP2, Ligne_Tableau_fils.VOI2 "
Sql = Sql & "FROM Ligne_Tableau_fils "
Sql = Sql & "WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & Me.Tag & "  "
Sql = Sql & "AND Ligne_Tableau_fils.ACTIVER=True AND  "
Sql = Sql & "(Ligne_Tableau_fils.APP='" & Replace("" & MyRange(I, 2), "*", ".") & "' Or [app2]='" & Replace("" & MyRange(I, 2), "*", ".") & "');"
Set Rs = Con.OpenRecordSet(Sql)
'If Rs.EOF = False Then
'260*AA.D3
'
'
'
'    Sql2 = "SELECT Ligne_Tableau_fils.APP, Max(Ligne_Tableau_fils.VOI) AS MaxDeVOI "
'    Sql2 = Sql2 & "FROM Ligne_Tableau_fils "
'    Sql2 = Sql2 & "WHERE Ligne_Tableau_fils.APP='" & Replace(MyRange(I, 2), "*", ".") & "'  "
'        Sql2 = Sql2 & "AND Ligne_Tableau_fils.ACTIVER=True  "
'    Sql2 = Sql2 & "AND Ligne_Tableau_fils.Id_IndiceProjet=" & Me.Tag & " "
'    Sql2 = Sql2 & "GROUP BY Ligne_Tableau_fils.APP, Ligne_Tableau_fils.ACTIVER, Ligne_Tableau_fils.Id_IndiceProjet;"
'    Set RsPasTrouver = Con.OpenRecordSet(Sql2)
'    If RsPasTrouver.EOF = Falset Then
'        MaxPin = RsPasTrouver!MaxDeVOI
'    End If
'
'    Sql2 = "SELECT Ligne_Tableau_fils.APP2, Max(Ligne_Tableau_fils.VOI2) AS MaxDeVOI "
'    Sql2 = Sql2 & "FROM Ligne_Tableau_fils "
'    Sql2 = Sql2 & "WHERE Ligne_Tableau_fils.APP2='" & Replace(MyRange(I, 2), "*", ".") & "'  "
'        Sql2 = Sql2 & "AND Ligne_Tableau_fils.ACTIVER=True  "
'    Sql2 = Sql2 & "AND Ligne_Tableau_fils.Id_IndiceProjet=" & Me.Tag & " "
'    Sql2 = Sql2 & "GROUP BY Ligne_Tableau_fils.APP2, Ligne_Tableau_fils.ACTIVER, Ligne_Tableau_fils.Id_IndiceProjet;"
'    Set RsPasTrouver = Con.OpenRecordSet(Sql2)
'    If RsPasTrouver.EOF = Falset Then
'        If MaxPin < RsPasTrouver!MaxDeVOI Then
'            MaxPin = RsPasTrouver!MaxDeVOI
'        End If
'    End If
'    Set RsPasTrouver = Con.CloseRecordSet(RsPasTrouver)
'    If Val("" & MaxPin) > 0 Then
'    For IEP = 1 To Val("" & MaxPin)
'        If VoieUne = False Then
'        VoieUne = True
'
'                MyRange(I + Ofset, 4) = ""
'                MyRange(I + Ofset, 5) = "" & IEP
'                MyRange(I + Ofset, 6) = MyRange(I, 2) & "." & MyRange(I + Ofset, 5)
'
'    Else
'        Ofset = Ofset + 1
'        InsertRow MyRange, I + Ofset
'
'                MyRange(I + Ofset, 4) = ""
'                MyRange(I + Ofset, 5) = "" & IEP
'                MyRange(I + Ofset, 6) = MyRange(I, 2) & "." & MyRange(I + Ofset, 5)
'
'    End If
'    Next
'    End If
'End If
While Rs.EOF = False
DoEvents

    If UCase(Trim("" & Rs!App)) = UCase(Replace(Trim("" & MyRange(I, 2)), "*", ".")) Then
        If ChercheXls(MyRange(I, 2) & "." & Rs!voi, MySheet.Range("a1").CurrentRegion, False, False, 2) = 0 Then
            Ofset = Ofset + 1
            InsertRow MyRange, I + Ofset
            MyRange(I + Ofset, 4).Select
            MyRange(I + Ofset, 4) = ""
             MyRange(I + Ofset, 5) = "" & Rs!voi
             MyRange(I + Ofset, 6) = MyRange(I, 2) & "." & MyRange(I + Ofset, 5)
         End If
   End If
   If UCase(Trim("" & Rs!app2)) = UCase(Replace(Trim("" & MyRange(I, 2)), "*", ".")) Then
         If ChercheXls(MyRange(I, 2) & "." & Rs!voi2, MySheet.Range("a1").CurrentRegion, False, False, 2) = 0 Then
            Ofset = Ofset + 1
            InsertRow MyRange, I + Ofset
            MyRange(I + Ofset, 4).Select
            MyRange(I + Ofset, 5) = "" & Rs!voi2
            MyRange(I + Ofset, 6) = MyRange(I, 2) & "." & MyRange(I + Ofset, 5)
            MyRange(I + Ofset, 4) = ""
         End If
    End If
    
    Rs.MoveNext
Wend

'End If
Next
MySheet2.Select
Set Myrange2 = MySheet2.Range("a1").CurrentRegion

'Sql = "SELECT T_Lien_Con_Famille.Connecteur, T_Lien_Con_Famille_Voies.Voie "
'Sql = Sql & "FROM T_Lien_Con_Famille INNER JOIN T_Lien_Con_Famille_Voies  "
'Sql = Sql & "ON T_Lien_Con_Famille.Id = T_Lien_Con_Famille_Voies.Id_T_Lien_Con_Famille "
For I = Myrange2.Rows.Count To 2 Step -1
'If I = 78 Then MsgBox ""
DoEvents
 VoieUne = False
 Ofset = 0
MySplit = Myrange2(I, 1).Value
Myrange2(I, 2).Select
Myrange2(I, 2) = Replace("" & Myrange2(I, 2), ".", "*")

Myrange2(I, 3) = Myrange2(I, 2)
MySplit = Split(MySplit & "§", "§")
'If UBound(MySplit) > 1 Then MsgBox ""

Myrange2(I, 1).Value = MySplit(0)
'If UCase(Myrange2(I, 4)) = "VRAI" Or UCase(Myrange2(I, 4)) = "YES" Then
    NbVoiEpisur = Val(Trim(Replace(Myrange2(I, 1), "EP", "")))
    NbVoiEpisur = NbVoiEpisur / 2
    Myrange2(I, 1) = "'IS-Term1"
    For IEP = 0 To 49
    If VoieUne = False Then
        VoieUne = True
        Myrange2(I + Ofset, 4) = ""
         Myrange2(I + Ofset, 5) = "P" & IEP
         Myrange2(I + Ofset, 4) = "P" & IEP 'Myrange2(I, 2) & "." & Myrange2(I + Ofset, 5)
         Ofset = Ofset + 1
'        InsertRow Myrange2, I + Ofset
'        Myrange2(I + Ofset, 5) = "D" & IEP
'        Myrange2(I + Ofset, 6) = Myrange2(I, 2) & "." & Myrange2(I + Ofset, 5)
    Else
        
        
        InsertRow Myrange2, I + Ofset
         Myrange2(I + Ofset, 4) = ""
         Myrange2(I + Ofset, 5) = "P" & IEP
         Myrange2(I + Ofset, 4) = "P" & IEP 'Myrange2(I, 2) & "." & Myrange2(I + Ofset, 5)
         Ofset = Ofset + 1
'        InsertRow Myrange2, I + Ofset
'        Myrange2(I + Ofset, 5) = "D" & IEP
'        Myrange2(I + Ofset, 6) = Myrange2(I, 2) & "." & Myrange2(I + Ofset, 5)
    End If
    Next
'Else
'SqlWher = "WHERE T_Lien_Con_Famille.Connecteur='" & MySplit(0) & "';"
'Set Rs = Con.OpenRecordSet(Sql & SqlWher)
'While Rs.EOF = False
'DoEvents
'    If VoieUne = False Then
'        VoieUne = True
'        Myrange2(I + Ofset, 4) = ""
'         Myrange2(I + Ofset, 5) = "" & Rs!voie
'         Myrange2(I + Ofset, 6) = Myrange2(I, 2) & "." & Myrange2(I + Ofset, 5)
'    Else
'        Ofset = Ofset + 1
'        InsertRow Myrange2, I + Ofset
'         Myrange2(I + Ofset, 5) = "" & Rs!voie
'         Myrange2(I + Ofset, 6) = Myrange2(I, 2) & "." & Myrange2(I + Ofset, 5)
'          Myrange2(I + Ofset, 4) = ""
'    End If
'
'    Rs.MoveNext
'Wend
'End If
Next




Sql = "SELECT Ligne_Tableau_fils.FIL, Ligne_Tableau_fils.LIAI, frm.txt28, '' AS Expr1, '' AS Expr2, Ligne_Tableau_fils.TEINT, Ligne_Tableau_fils.APP, Ligne_Tableau_fils.VOI, Ligne_Tableau_fils.APP2, Ligne_Tableau_fils.VOI2 "
Sql = Sql & "FROM Ligne_Tableau_fils LEFT JOIN (SELECT con_contacts.*, lst31.*, lst29.* "
Sql = Sql & "FROM (con_contacts LEFT JOIN lst31 ON con_contacts.lst31 = lst31.CatID) LEFT JOIN lst29 ON con_contacts.lst29 =  "
Sql = Sql & "lst29.CatID in '"
Sql = Sql & TableauPath("Eb_FILS")
Sql = Sql & "') AS frm ON (Ligne_Tableau_fils.SECT = frm.lst29.CatName2) AND (Ligne_Tableau_fils.ISO = frm.lst31.CatName2) "
Sql = Sql & "WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & Me.Tag & "  "
Sql = Sql & "AND Ligne_Tableau_fils.ACTIVER=True "
Sql = Sql & "GROUP BY Ligne_Tableau_fils.FIL, Ligne_Tableau_fils.LIAI,  Max(frm.txt28) AS MaxDetxt28, '', '', Ligne_Tableau_fils.TEINT, Ligne_Tableau_fils.APP, Ligne_Tableau_fils.VOI, Ligne_Tableau_fils.APP2, Ligne_Tableau_fils.VOI2, Ligne_Tableau_fils.Id_IndiceProjet, Ligne_Tableau_fils.ACTIVER;"

Sql = "SELECT Ligne_Tableau_fils.FIL, Ligne_Tableau_fils.LIAI,  Max(frm.txt28) AS MaxDetxt28, '' AS Expr1, '' AS Expr2, Ligne_Tableau_fils.TEINT,  "
Sql = Sql & "Ligne_Tableau_fils.APP, Ligne_Tableau_fils.VOI, Ligne_Tableau_fils.APP2, Ligne_Tableau_fils.VOI2 "
Sql = Sql & "FROM ((Ligne_Tableau_fils LEFT JOIN (SELECT con_contacts.*, lst31.*, lst29.* FROM (con_contacts LEFT JOIN lst31  "
Sql = Sql & "ON con_contacts.lst31 = lst31.CatID) LEFT JOIN lst29 ON con_contacts.lst29 =  lst29.CatID in '"
Sql = Sql & TableauPath("Eb_FILS")
Sql = Sql & "') AS frm ON (Ligne_Tableau_fils.ISO = frm.lst31.CatName2) AND (Ligne_Tableau_fils.SECT = frm.lst29.CatName2))  "
Sql = Sql & "LEFT JOIN T_Critères ON (Ligne_Tableau_fils.OPTION = T_Critères.CRITERES) AND (Ligne_Tableau_fils.Id_IndiceProjet =  "
Sql = Sql & "T_Critères.Id_IndiceProjet)) LEFT JOIN T_Critères AS T_Critères_1 ON (Ligne_Tableau_fils.OPTION = T_Critères_1.CODE_CRITERE)  "
Sql = Sql & "AND (Ligne_Tableau_fils.Id_IndiceProjet = T_Critères_1.Id_IndiceProjet) "


Sql = Sql & "Where Ligne_Tableau_fils.Id_IndiceProjet = " & Me.Tag & " And Ligne_Tableau_fils.ACTIVER = True "
Sql = Sql & " and  ((T_Critères.CRITERES Is Null AND T_Critères_1.CRITERES Is Null)  "
Sql = Sql & "OR ';' & T_Critères.CRITERES & ';'  Like '%;tous;%' "
Sql = Sql & "OR ';' & T_Critères_1.CRITERES & ';'  Like '%;tous;%' "
For I = 0 To UBound(SplitEquipement) - 1
    If Trim("" & SplitEquipement(I)) <> "" Then
    MySplit2 = Split(SplitEquipement(I) & "_", "_")
        Sql = Sql & "OR ';' & T_Critères_1.CRITERES & ';'  Like '%;" & MySplit2(0) & ";%' "
        Sql = Sql & "OR ';' & T_Critères.CRITERES & ';'  Like '%;" & MySplit2(0) & ";%' "
    End If
Next

'sql = sql & "( OR (((T_Critères_1.CRITERES)='GRAND FROID'))"

Sql = Sql & ") GROUP BY Ligne_Tableau_fils.FIL, Ligne_Tableau_fils.LIAI,  '', '', Ligne_Tableau_fils.TEINT,  "
Sql = Sql & "Ligne_Tableau_fils.APP, Ligne_Tableau_fils.VOI, Ligne_Tableau_fils.APP2, Ligne_Tableau_fils.VOI2,  "
Sql = Sql & "Ligne_Tableau_fils.Id_IndiceProjet, Ligne_Tableau_fils.ACTIVER, T_Critères.CRITERES, T_Critères_1.CRITERES ;"


Sql = "SELECT Ligne_Tableau_fils.FIL, Ligne_Tableau_fils.LIAI,  Max(frm.txt28) AS MaxDetxt28, '' AS Expr1, '' AS Expr2, Ligne_Tableau_fils.TEINT,  "
Sql = Sql & "Ligne_Tableau_fils.APP, Ligne_Tableau_fils.VOI, Ligne_Tableau_fils.APP2, Ligne_Tableau_fils.VOI2 "
Sql = Sql & "FROM Ligne_Tableau_fils LEFT JOIN (SELECT con_contacts.*, lst31.*, lst29.* FROM (con_contacts LEFT JOIN lst31   "
Sql = Sql & "ON con_contacts.lst31 = lst31.CatID) LEFT JOIN lst29 ON con_contacts.lst29 =  lst29.CatID in '"
Sql = Sql & TableauPath("Eb_FILS")
Sql = Sql & "') AS frm ON (Ligne_Tableau_fils.SECT = frm.lst29.CatName2) AND (Ligne_Tableau_fils.ISO = frm.lst31.CatName2) "

Sql = Sql & "Where Ligne_Tableau_fils.Id_IndiceProjet = " & Me.Tag & "  "
Sql = Sql & "And Ligne_Tableau_fils.ACTIVER = True "

Sql = Sql & "and ( ';' & Ligne_Tableau_fils.OPTION & ';'  Like '%;tous;%' "
Sql = Sql & "OR ';' & Ligne_Tableau_fils.OPTION & ';'  Like '%;ALL;%' "

For I = 0 To UBound(SplitEquipement) - 1
    If Trim("" & SplitEquipement(I)) <> "" Then
    MySplit2 = Split(SplitEquipement(I) & "_", "_")
          Sql = Sql & "OR ';' & Ligne_Tableau_fils.OPTION & ';'  Like '%;" & MySplit2(0) & ";%' "
    End If
Next

'Sql = Sql & ") GROUP  Ligne_Tableau_fils.FIL, Ligne_Tableau_fils.LIAI, frm.txt28, '' AS Expr1, '' AS Expr2, Ligne_Tableau_fils.TEINT,  "
'Sql = Sql & "Ligne_Tableau_fils.APP, Ligne_Tableau_fils.VOI, Ligne_Tableau_fils.APP2, Ligne_Tableau_fils.VOI2, Ligne_Tableau_fils.OPTION ;"
'Sql = Sql & "WHERE (((Ligne_Tableau_fils.Id_IndiceProjet)=1342) AND ((Ligne_Tableau_fils.ACTIVER)=True)) OR (((Ligne_Tableau_fils.Id_IndiceProjet)=1342) AND ((Ligne_Tableau_fils.ACTIVER)=True) AND ((';' & [Ligne_Tableau_fils].[OPTION] & ';') Like '*;tous;*')) OR (((Ligne_Tableau_fils.Id_IndiceProjet)=1342) AND ((Ligne_Tableau_fils.ACTIVER)=True)) OR (((Ligne_Tableau_fils.Id_IndiceProjet)=1342) AND ((Ligne_Tableau_fils.ACTIVER)=True)) OR (((Ligne_Tableau_fils.Id_IndiceProjet)=1342) AND ((Ligne_Tableau_fils.ACTIVER)=True) AND ((';' & [Ligne_Tableau_fils].[OPTION] & ';') Like '*;;*'));"


Sql = Sql & ") GROUP BY Ligne_Tableau_fils.FIL, Ligne_Tableau_fils.LIAI,  '', '', Ligne_Tableau_fils.TEINT,  "
Sql = Sql & "Ligne_Tableau_fils.APP, Ligne_Tableau_fils.VOI, Ligne_Tableau_fils.APP2, Ligne_Tableau_fils.VOI2  "


'HAVING ;



Set Rs = Con.OpenRecordSet(Sql)
Set MySheet = MyClaseur.Sheets("Wire")
MySheet.Select
MySheet.Range("A2").CopyFromRecordset Rs
Set MyRange = MySheet.Range("A1").CurrentRegion

For I = 2 To MyRange.Rows.Count
'If I = 150 Then MsgBox ""
MyRange(I, 7).Select
MyVal = CDbl(Trim(Replace("0" & MyRange(I, 3), "mm", "")))
MyRange(I, 3) = MyVal / 1000 & "mm"
txtSheetName = "SIC-TERM"
SershRow = SershXls(MyClaseur.Sheets("SIC-TERM"), Replace(MyRange(I, 7), ".", "*") & "." & MyRange(I, 8))
SershRow = SershXls(MyClaseur.Sheets("SIC-TERM"), Replace(MyRange(I, 7), ".", "*"))
If SershRow <> 0 Then

'   1342  If Trim("" & MyClaseur.Sheets("SIC-TERM").Cells(SershRow, 1)) = "" Then
'        MyRange(I, 7) = Replace(MyRange(I, 7), ".", "*") & "." & MyRange(I, 8)
'
'     Else
'        MyRange(I, 7) = Replace(MyRange(I, 7), ".", "*")
'
'
'     End If
MyRange(I, 7) = Replace(MyRange(I, 7), ".", "*") & "." & MyRange(I, 8)
    Else
        SershRow = SershXls(MyClaseur.Sheets("SIC-TERM"), Replace(MyRange(I, 7), ".", "*"))
        If SershRow <> 0 Then
           If MyClaseur.Sheets("SIC-TERM").Cells(SershRow, 1) = "IS-Term1" Then
            PnG = 1
            PnD = 2
            PnT = 0
           For Ofset = 0 To 49
           If MyRange(I, 8) = "G" Then
            PnT = PnG
           Else
            If MyRange(I, 8) = "D" Then
                PnT = PnD
            Else
                PnT = Ofset
            End If
           End If
                If Trim("" & MyClaseur.Sheets("SIC-TERM").Cells(SershRow + PnT, 100)) = "" Then
                    MyClaseur.Sheets("SIC-TERM").Cells(SershRow + PnT, 100) = 1
                     If Trim("" & MyClaseur.Sheets("SIC-TERM").Cells(SershRow + PnT, 1)) <> "" Then
                        MyRange(I, 7) = Replace(MyRange(I, 7), ".", "*")
                     Else
                        MyRange(I, 7) = Replace(MyRange(I, 7), ".", "*") & "." & MyClaseur.Sheets("SIC-TERM").Cells(SershRow + PnT, 5)
                    End If
                    
                    Exit For
                End If
                PnG = PnG + 2
                    PnD = PnG + 2
           Next
          Else
            MyRange(I, 7) = Replace(MyRange(I, 7), ".", "*")
            If InStr(1, UCase(MyRange(I, 8)), UCase("cn")) <> 0 Then
                MyRange(I, 7) = MyRange(I, 7) & "_" & MyRange(I, 8)
               MyClaseur.Sheets("IS").Cells(MyClaseur.Sheets("IS").Range("a1").CurrentRegion.Rows.Count + 1, 1) = "'IS-Term1"
               MyClaseur.Sheets("IS").Cells(MyClaseur.Sheets("IS").Range("a1").CurrentRegion.Rows.Count, 2) = "'" & MyRange(I, 7)
               MyClaseur.Sheets("IS").Cells(MyClaseur.Sheets("IS").Range("a1").CurrentRegion.Rows.Count, 3) = "'" & MyRange(I, 7)
               MyClaseur.Sheets("IS").Cells(MyClaseur.Sheets("IS").Range("a1").CurrentRegion.Rows.Count, 4) = "'P0"
               MyClaseur.Sheets("IS").Cells(MyClaseur.Sheets("IS").Range("a1").CurrentRegion.Rows.Count, 5) = "'P0"
                  MyRange(I, 7) = MyRange(I, 7) & ".P0"
            End If
          End If
        Else
        MyRange(I, 7) = Replace(MyRange(I, 7), ".", "*")
        End If
       
    End If
        
SershRow = SershXls(MyClaseur.Sheets("SIC-TERM"), Replace(MyRange(I, 9), ".", "*") & "." & MyRange(I, 10))
SershRow = SershXls(MyClaseur.Sheets("SIC-TERM"), Replace(MyRange(I, 9), ".", "*"))
If SershRow <> 0 Then
'    If Trim("" & MyClaseur.Sheets("SIC-TERM").Cells(SershRow, 1)) = "" Then
'                MyRange(I, 8) = Replace(MyRange(I, 9), ".", "*") & "." & MyRange(I, 10)
'
'     Else
'
'         MyRange(I, 8) = Replace(MyRange(I, 9), ".", "*")
'
'     End If
     MyRange(I, 8) = Replace(MyRange(I, 9), ".", "*") & "." & MyRange(I, 10)
    Else
       
       SershRow = SershXls(MyClaseur.Sheets("IS"), Replace(MyRange(I, 9), ".", "*"))
        If SershRow <> 0 Then
           If MyClaseur.Sheets("is").Cells(SershRow, 1) = "IS-Term1" Then
            PnG = 1
            PnD = 2
            PnT = 0
           For Ofset = 0 To 49
           If MyRange(I, 10) = "G" Then
            PnT = PnG
           Else
            If MyRange(I, 10) = "D" Then
                PnT = PnD
            Else
                PnT = Ofset
            End If
           End If
                If Trim("" & MyClaseur.Sheets("is").Cells(SershRow + PnT, 100)) = "" Then
                    MyClaseur.Sheets("is").Cells(SershRow + PnT, 100) = 1
                     If Trim("" & MyClaseur.Sheets("is").Cells(SershRow + PnT, 1)) <> "" Then
                        MyRange(I, 8) = Replace(MyRange(I, 9), ".", "*")
                     Else
                        MyRange(I, 8) = Replace(MyRange(I, 9), ".", "*") & "." & MyClaseur.Sheets("is").Cells(SershRow + PnT, 5)
                    End If
                    
                    Exit For
                End If
                PnG = PnG + 2
                    PnD = PnG + 2
           Next
          Else
            MyRange(I, 7) = Replace(MyRange(I, 8), ".", "*")
            If InStr(1, UCase(MyRange(I, 8)), UCase("cn")) <> 0 Then
                MyRange(I, 7) = MyRange(I, 7) & "_" & MyRange(I, 8)
               MyClaseur.Sheets("IS").Cells(MyClaseur.Sheets("IS").Range("a1").CurrentRegion.Rows.Count + 1, 1) = "'IS-Term1"
               MyClaseur.Sheets("IS").Cells(MyClaseur.Sheets("IS").Range("a1").CurrentRegion.Rows.Count, 2) = "'" & MyRange(I, 7)
               MyClaseur.Sheets("IS").Cells(MyClaseur.Sheets("IS").Range("a1").CurrentRegion.Rows.Count, 3) = "'" & MyRange(I, 7)
               MyClaseur.Sheets("IS").Cells(MyClaseur.Sheets("IS").Range("a1").CurrentRegion.Rows.Count, 4) = "'P0"
               MyClaseur.Sheets("IS").Cells(MyClaseur.Sheets("IS").Range("a1").CurrentRegion.Rows.Count, 5) = "'P0"
                  MyRange(I, 7) = MyRange(I, 7) & ".P0"
            End If
          End If
        Else
        MyRange(I, 7) = Replace(MyRange(I, 8), ".", "*")
        End If
        
    End If
     MyRange(I, 9) = ""
        MyRange(I, 10) = ""
        
Next
'CheckI MyClaseur

Set MySheet = MyClaseur.Sheets("Create")
MySheet.Select
'Set Myrange = MySheet.Range("A1").CurrentRegion
'For I = 2 To Myrange.Rows.Count
'Myrange(I, 1).Select
'If Trim("" & Myrange(I, 1)) <> "" Then
' PnG = -1
' PnD = 0
'    If Myrange(I, 1) = "IS-Term1" Then
'        IS_Term1 = True
'        ValEpisure = Myrange(I, 2)
'    Else
'        IS_Term1 = False
'
'    End If
'End If
'    If IS_Term1 = True Then
'        If Left(UCase(Myrange(I, 5)), 1) = "G" Then
'            PnG = PnG + 2
'            Myrange(I, 5) = "P" & PnG
'        Else
'            PnD = PnD + 2
'                Myrange(I, 5) = "P" & PnD
'        End If
'        DoEvents
'        ReplaceTousXls MyClaseur.Sheets("Wire"), Myrange(I, 6), ValEpisure & "." & Myrange(I, 5)
'    End If
'Next
Set MySheet = MyClaseur.Sheets("SIC-TERM")
Set MyRange = MySheet.Range("a1").CurrentRegion
Set MyRange = MySheet.Range("E2:E" & MyRange.Rows.Count)
MyRange.Replace "FAUX", ""
MyRange.Replace "FALSE", ""
Create PathPl, "xml", PathPl, MyClaseur
MyClaseur.Sheets("Create").Select
MyClaseur.SaveAs PathPl
MyClaseur.Close False

MyExcel.Quit

Else
NomenclatureOk = True
 UserForm2.chargement txt6, CLng(Me.txt3.Tag), txt9, Me, Edition:=True

End If
MsgBox "Fin de traitement", vbInformation
End Sub
Private Sub CommandButton3_Click()

Noquite = False
Unload Me


End Sub
Public Sub Charge(MyForm As Object)
Dim Sql As String
Dim Rs As Recordset
IdIndiceProjet = MyForm.IdIndiceProjet

Sql = "SELECT SelectProjets.* FROM SelectProjets WHERE SelectProjets.Id=" & IdIndiceProjet & " ;"

Set Rs = Con.OpenRecordSet(Sql)

Set FormBarGrah = Me
If Rs.EOF = False Then
For I = 0 To 11
    Me.Controls("txt" & CStr(I + 1)) = "" & Rs(I)
     Me.Controls("txt" & CStr(I + 1)).Tag = "" & Rs.Fields(12)

Next I
    
    OptionButton2.Value = True
    OptionButton1.Value = False
    Me.CommandButton1.Enabled = True
 End If
 Set Rs = Con.CloseRecordSet(Rs)
 MyForm.Hide
 Me.Show
End Sub

Private Sub OptionButton1_Click()
OptionButton2.Value = False
End Sub

Private Sub OptionButton2_Click()
OptionButton1.Value = False
End Sub

Private Sub UserForm_Activate()
Noquite = True
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Cancel = Noquite
Con.CloseConnection
End Sub
Sub CalculIndexPin(Valeure As String, AvantValeure As String, Mytype As String, Index As Long)

If IsNumeric("" & Valeur) = True Then
    Mytype = "1"
Else
End If
Select Case UCase(Mytype)

'        1,2,3...
        Case "1"
               If Val("" & Valeure) < Val("" & AvantValeure) Then Index = Valeur
                
'        10,9,8..
        Case "10"
                
                
                
'        1-1,1-2,1-3...
         Case "11"
                    
'        10-1,10-2,10-3...
        Case "101"
                  
'        1-10,1-9,1-8...
        Case "110"
           
                    
'        10-10,10-9,10-8...
        Case "1010"
            
                    
'        A,B,C...
        Case "A"
           
'        A1,A2,A3...
         Case "A1"
                    
'        A10,A9,A8...
        Case "A10"
       
'       Z,Y,X...
       Case "Z"
'        A-A,A-B,A-C...
        Case "AA"
                    
'        A-Z,A-Y,A-X...
        Case "AZ"
                    
'        Z-A,Z-B,Z-C...
        Case "ZA"
                    
'        Z-Z,Z-Y,Z-X...
        Case "ZZ"
       
'        1A,1B,1C...
        Case "1A"
        
'        10A,10B,10C...
        Case "10A"

'        Z1,Z2,Z3...
        Case "Z1"
                    
'        Z10,Z9,Z8...
        Case "Z10"
                    
'        1Z,1Y,1Z...
        Case "1Z"
                    
'        10Z,10Y,10X...
        Case "10Z"
       

       
        Case Else

End Select
End Sub
