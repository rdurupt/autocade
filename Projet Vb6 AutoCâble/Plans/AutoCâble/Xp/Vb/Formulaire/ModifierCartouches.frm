VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ModifierCartouches 
   Caption         =   "Modifier le cartouche :"
   ClientHeight    =   10125
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12300
   Icon            =   "ModifierCartouches.dsx":0000
   OleObjectBlob   =   "ModifierCartouches.dsx":030A
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ModifierCartouches"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public IdProjet As Long
Public IdPieces As Long
Public IdIndiceProjet As Long
Dim Noquite As Boolean
Public Execute As Boolean
Dim ChronoAnneeEnCours As String
Dim ChronoAnnee_M1 As String
Dim ChronoAnneeEnCoursOld As String



Private Sub CommandButton1_Click()
Dim Sql As String
Dim Rs As Recordset
Dim RsListe As Recordset

Sql = "SELECT T_Liste_Projet.Projet FROM T_Liste_Projet ORDER BY T_Liste_Projet.Projet;"
Set RsListe = Con.OpenRecordSet(Sql)
txt1.Clear
txt24.Clear



Dim CherchPicesAnnuler As Boolean
If Trim("" & BdDateTable) <> "" Then
    RqChronoAnne = "[Chrono Requête " & BdDateTable & "]"
    ChronoAnneeEnCours = "[Chrono" & BdDateTable & "]"
    ChronoAnnee_M1 = "[Chrono" & Val(BdDateTable) - 1 & "]"
    ChronoAnneeEnCoursOld = "[Chrono" & BdDateTable & "Old]"
Else
     RqChronoAnne = "[Chrono Requête " & Format(Date, "yyyy]")
     ChronoAnneeEnCours = "[Chrono" & Format(Date, "yyyy]")
     ChronoAnneeEnCoursOld = "[Chrono" & Format(Date, "yyyy") & "Old]"
     ChronoAnnee_M1 = "[Chrono" & Val(Format(Date, "yyyy")) - 1 & "]"
End If

CherchPicesAnnuler = False
CherchPices.Charge Me, " LiAutoCadSave <>  Null and IdStatus<3 and Archiver=False", True
CherchPicesAnnuler = CherchPices.Annuler
Unload CherchPices
If CherchPicesAnnuler = True Then Exit Sub
IdFils = 0
Sql = "SELECT T_indiceProjet.Pere FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Me.Tag & ";"
Set Rs = Con.OpenRecordSet(Sql)
If Rs!Pere > 0 Then Me.Tag = Rs!Pere

Sql = "SELECT  RqCartouche.Ref_PF, RqCartouche.Ref_Plan_CLI, RqCartouche.Ref_Piece_CLI,RqCartouche.Projet AS txt1,  "
Sql = Sql & "RqCartouche.Vague AS txt2,  "
Sql = Sql & "RqCartouche.Equipement AS txt3,  "
Sql = Sql & "RqCartouche.Responsable AS txt4,  "
Sql = Sql & "RqCartouche.Ensemble AS txt5,  "
Sql = Sql & "RqCartouche.CleAc AS txt6,  "
Sql = Sql & "RqCartouche.PI & '_' & RqCartouche.PI_Indice AS txt7,  "
Sql = Sql & "RqCartouche.PL & '_' & RqCartouche.PL_Indice AS txt8,  "
Sql = Sql & "RqCartouche.[OU] & '_' & RqCartouche.OU_Indice AS txt9,  "
Sql = Sql & "RqCartouche.Li & '_' &  RqCartouche.LI_Indice AS txt10,  "
Sql = Sql & "RqCartouche.Client AS txt11,  "
Sql = Sql & "RqCartouche.Destinataire AS txt12,  "
Sql = Sql & "RqCartouche.Service AS txt13,  "
Sql = Sql & "RqCartouche.RefPF AS txt14, "
Sql = Sql & " RqCartouche.RefP AS txt15,  "
Sql = Sql & "RqCartouche.DessineDate AS txt16,  "
Sql = Sql & "RqCartouche.DessineNOM AS txt17,  "
Sql = Sql & "RqCartouche.VerifieDate AS txt18,  "
Sql = Sql & "RqCartouche.VerifieNom AS txt19,  "
Sql = Sql & "RqCartouche.ApprouveDate AS txt20,  "
Sql = Sql & "RqCartouche.ApprouveNom AS txt21, "
Sql = Sql & "RqCartouche.NbCartouche AS txt22, "
Sql = Sql & "RqCartouche.RefPieceClient AS txt23, "
Sql = Sql & "RqCartouche.BaseVehicule AS txt24, "
Sql = Sql & "RqCartouche.Masse AS txt25 "

Sql = Sql & "FROM RqCartouche "
Sql = Sql & "WHERE RqCartouche.T_indiceProjet.Id=" & Val(Me.Tag) & " ;"
Debug.Print Sql
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = False Then
 Me.Ref_PF = "" & Rs!Ref_PF
  Me.Ref_Plan_CLI = "" & Rs!Ref_Plan_CLI
  Me.Ref_Piece_CLI = "" & Rs!Ref_Piece_CLI
' Me.Controls("txt" & CStr(1)) = "" & Rs.Fields("txt" & CStr(1))
txt1.Clear

 While RsListe.EOF = False
   
    txt1.AddItem Trim("" & RsListe!Projet)
     If Trim("" & RsListe!Projet) = Trim("" & Rs.Fields("txt" & CStr(1))) Then
        txt1.ListIndex = txt1.ListCount - 1
     End If
    RsListe.MoveNext
Wend
RsListe.Requery
txt24.Clear
txt24.AddItem ""
 While RsListe.EOF = False
   
    txt24.AddItem Trim("" & RsListe!Projet)
     If Trim("" & RsListe!Projet) = Trim("" & Rs.Fields("txt24")) Then
        txt24.ListIndex = txt24.ListCount - 1
    End If
    RsListe.MoveNext
Wend
TXT25 = Trim("" & Rs.Fields("txt25"))
Set RsListe = Con.CloseRecordSet(RsListe)
  Me.Controls("txt1Bis") = "" & Rs.Fields("txt" & CStr(1))
For i = 2 To 3
    Me.Controls("txt" & CStr(i)) = "" & Rs.Fields("txt" & CStr(i))
Next i
 Me.Controls("txt" & CStr(4)) = "" & Rs.Fields("txt" & CStr(4))
  Me.Controls("txt" & CStr(5)) = "" & Rs.Fields("txt" & CStr(5))
For i = 6 To 12
    Me.Controls("txt" & CStr(i)) = "" & Rs.Fields("txt" & CStr(i))
Next i
For i = 13 To 15
    Me.Controls("txt" & CStr(i)) = "" & Rs.Fields("txt" & CStr(i))
Next i
For i = 16 To 18 Step 2
    Me.Controls("txt" & CStr(i)) = "" & Rs.Fields("txt" & CStr(i))
Next i
 Me.Controls("txt" & CStr(20)) = "" & Rs.Fields("txt" & CStr(20))
For i = 17 To 21 Step 2
    Me.Controls("txt" & CStr(i)) = "" & Rs.Fields("txt" & CStr(i))
Next i
 Me.Controls("txt" & CStr(22)) = "" & Rs.Fields("txt" & CStr(22))
  Me.Controls("txt" & CStr(23)) = "" & Rs.Fields("txt" & CStr(23))
End If
InitLPlanche Val(Me.Tag)
CleCh = Split(Me.txt7, "_")
If ConBaseNum.OpenConnetion(DbNumPlan) = True Then

Sql = "SELECT " & ChronoAnneeEnCours & ".Destinataire,Agent_2.[Nom ag] AS Red_Nom, Agent_2.[Prénom ag] AS Red_P_Nom,  "
Sql = Sql & "Agent_1.[Nom ag] AS Verif_Nom, Agent_1.[Prénom ag] AS Verif_P_Nom,  "
Sql = Sql & "Agent.[Nom ag] AS Apr_Nom,   "
Sql = Sql & "Agent.[Prénom ag] AS Apr_P_Nom " ', [Clé ty] & '_' "
'Sql = Sql & "& [Clé ac] & '_ ' & [Année] & '_' & [Clé Ch] & '_' & [Rév] AS PI  "
Sql = Sql & "FROM ((" & ChronoAnneeEnCours & " INNER JOIN Agent ON " & ChronoAnneeEnCours & ".[Clé ap] = Agent.[Clé ag])   "
Sql = Sql & "INNER JOIN Agent AS Agent_1 ON " & ChronoAnneeEnCours & ".[Clé ve] = Agent_1.[Clé ag])   "
Sql = Sql & "INNER JOIN Agent AS Agent_2 ON " & ChronoAnneeEnCours & ".[Clé re] = Agent_2.[Clé ag]  "
Sql = Sql & "WHERE [Clé Ch] =" & CleCh(3) & " ;"

Set Rs = ConBaseNum.OpenRecordSet(Sql)
If Rs.EOF = True Then
    
    Sql = "SELECT " & ChronoAnnee_M1 & ".Destinataire,Agent_2.[Nom ag] AS Red_Nom, Agent_2.[Prénom ag] AS Red_P_Nom,  "
    Sql = Sql & "Agent_1.[Nom ag] AS Verif_Nom, Agent_1.[Prénom ag] AS Verif_P_Nom,  "
    Sql = Sql & "Agent.[Nom ag] AS Apr_Nom,   "
    Sql = Sql & "Agent.[Prénom ag] AS Apr_P_Nom " ', [Clé ty] & '_' "
'    Sql = Sql & "& [Clé ac] & '_ ' & [Année] & '_' & [Clé Ch] & '_' & [Rév] AS PI  "
    Sql = Sql & "FROM ((" & ChronoAnnee_M1 & " INNER JOIN Agent ON " & ChronoAnnee_M1 & ".[Clé ap] = Agent.[Clé ag])   "
    Sql = Sql & "INNER JOIN Agent AS Agent_1 ON " & ChronoAnnee_M1 & ".[Clé ve] = Agent_1.[Clé ag])   "
    Sql = Sql & "INNER JOIN Agent AS Agent_2 ON " & ChronoAnnee_M1 & ".[Clé re] = Agent_2.[Clé ag]  "
    Sql = Sql & "WHERE [Clé Ch] =" & CleCh(3) & " ;"


    Set Rs = ConBaseNum.OpenRecordSet(Sql)
    If Rs.EOF = True Then
       
        Sql = "SELECT " & ChronoAnneeEnCoursOld & ".Destinataire,Agent_2.[Nom ag] AS Red_Nom, Agent_2.[Prénom ag] AS Red_P_Nom,  "
        Sql = Sql & "Agent_1.[Nom ag] AS Verif_Nom, Agent_1.[Prénom ag] AS Verif_P_Nom,  "
        Sql = Sql & "Agent.[Nom ag] AS Apr_Nom,   "
        Sql = Sql & "Agent.[Prénom ag] AS Apr_P_Nom, [Clé ty] & '_' "
        Sql = Sql & "& [Clé ac] & '_ ' & [Année] & '_' & [Clé Ch] & '_' & [Rév] AS PI  "
        Sql = Sql & "FROM ((" & ChronoAnneeEnCoursOld & " INNER JOIN Agent ON " & ChronoAnneeEnCoursOld & ".[Clé ap] = Agent.[Clé ag])   "
        Sql = Sql & "INNER JOIN Agent AS Agent_1 ON " & ChronoAnneeEnCoursOld & ".[Clé ve] = Agent_1.[Clé ag])   "
        Sql = Sql & "INNER JOIN Agent AS Agent_2 ON " & ChronoAnneeEnCoursOld & ".[Clé re] = Agent_2.[Clé ag]  "
       Sql = Sql & "WHERE [Clé Ch] =" & CleCh(3) & " ;"

        Set Rs = ConBaseNum.OpenRecordSet(Sql)
    End If
End If
If Rs.EOF = False Then
    Red_Nom = Trim("" & Rs!Red_Nom)
    Red_P_Nom = Trim("" & Rs!Red_P_Nom)
    Verif_Nom = Trim("" & Rs!Verif_Nom)
    Verif_P_Nom = Trim("" & Rs!Verif_P_Nom)
    Apr_Nom = Trim("" & Rs!Apr_Nom)
    Apr_P_Nom = Trim("" & Rs!Apr_P_Nom)
    Destinataire = Trim("" & Rs!Destinataire)
     If Len(Destinataire) > 0 Then
     txt12 = UCase(Destinataire)
     End If
     
    If Len(Red_Nom) > 0 Then
        If Len(Red_P_Nom) > 0 Then
            txt17 = UCase(Red_Nom) & "." & UCase(Left(Red_P_Nom, 1))
        Else
           txt17 = UCase(Red_Nom)
        End If
    Else
        If Len(Red_P_Nom) > 0 Then
            txt17 = UCase(Red_P_Nom)
        End If
    End If
    
    
     If Len(Verif_Nom) > 0 Then
        If Len(Verif_P_Nom) > 0 Then
            txt19 = UCase(Verif_Nom) & "." & UCase(Left(Verif_P_Nom, 1))
        Else
           txt19 = UCase(Verif_Nom)
        End If
    Else
        If Len(Verif_P_Nom) > 0 Then
            txt19 = UCase(Verif_P_Nom)
        End If
    End If
    
    
    If Len(Apr_Nom) > 0 Then
        If Len(Apr_P_Nom) > 0 Then
            txt21 = UCase(Apr_Nom) & "." & UCase(Left(Apr_P_Nom, 1))
        Else
           txt21 = UCase(Apr_Nom)
        End If
    Else
        If Len(Apr_P_Nom) > 0 Then
            txt21 = UCase(Apr_P_Nom)
        End If
    End If
End If
Set Rs = ConBaseNum.CloseRecordSet(Rs)
ConBaseNum.CloseConnection
End If
End Sub
Sub InitLPlanche(Id_Pieces As Long)
Dim Rs As Recordset
Dim Sql As String
Sql = "SELECT T_indiceProjet.Cartouche "
Sql = Sql & " FROM T_indiceProjet "
Sql = Sql & " WHERE T_indiceProjet.Id=" & Id_Pieces & ";"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = False Then
    aa = "" & Rs!Cartouche
    If Trim(aa) <> "" Then
    aa = Split(aa, "\")
    MuPlanche = aa(UBound(aa))
    For i = 0 To PlanchClous.ListCount - 1
    If UCase(PlanchClous.List(i)) = UCase(MuPlanche) Then PlanchClous.ListIndex = i
    Next
    End If
End If
Set Rs = Con.CloseRecordSet(Rs)
End Sub

Private Sub CommandButton3_Click()


UserForm1.Charger txt5, vbCrLf, "Ensemble:"


End Sub


Private Sub CommandButton4_Click()


UserForm1.Charger txt3, ";", "Equipement:", "_"


End Sub

Private Sub CommandButton5_Click()

UserForm1.Charger txt2, " ", "Vagues:", " "


End Sub

Private Sub CommandButton7_Click()
Dim Sql As String
Dim Rs As Recordset
Dim MsgAutoCad As String
MsgAutoCad = ""
Set FormBarGrah = Me
If MyFormat("DATE", txt16, "Déssiné par") = False Then Exit Sub
If MyFormat("DATE", txt18, "Vérifié par") = False Then Exit Sub
If MyFormat("DATE", txt20, "Approuvé par") = False Then Exit Sub
If MyFormat("DBL", TXT25, "Masse") = False Then Exit Sub
If MyFormatQRY(txt22) = False Then Exit Sub
If Trim("" & Me.Tag) = "" Then
    CommandButton1_Click
    Exit Sub
End If
If Trim(PlanchClous.Text) = "" Then
    MsgBox "Vous devez sélectionner une planche à clous", vbExclamation
    Me.PlanchClous.SetFocus
    Exit Sub
End If

If txt1Bis.Value <> txt1.Value Then
Sql = "SELECT T_Projet.id  FROM T_Projet "
Sql = Sql & "WHERE T_Projet.Projet='" & MyReplace(Me.txt1) & "';"

Set Rs = Con.OpenRecordSet(Sql)
    If Rs.EOF = True Then
        Sql = "INSERT INTO T_Projet ( Projet ) VALUES ('" & MyReplace(Me.txt1) & "');"
        Con.Exequte Sql
        Rs.Requery
    End If
Sql = "UPDATE T_Pieces INNER JOIN T_indiceProjet ON T_Pieces.Id = T_indiceProjet.Id_Pieces  "
Sql = Sql & "SET T_Pieces.IdProjet = " & Rs!Id & " "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Me.Tag & "  "
Sql = Sql & "OR T_indiceProjet.Pere=" & Me.Tag & ";"
Con.Exequte Sql
End If
Sql = "UPDATE RqCartouche SET "
Sql = Sql & "RqCartouche.Projet = '" & MyReplace(txt1) & "', "
Sql = Sql & "RqCartouche.Vague = '" & MyReplace(txt2) & "', "
Sql = Sql & "RqCartouche.Equipement = '" & MyReplace(txt3) & "', "
Sql = Sql & "RqCartouche.Responsable = '" & MyReplace(TXT4) & "', "
Sql = Sql & "RqCartouche.Ensemble = '" & MyReplace(txt5) & "', "
Sql = Sql & "RqCartouche.CleAc = " & txt6 & ", "
Sql = Sql & "RqCartouche.Ref_PF = '" & UCase(MyReplace(Ref_PF)) & "', "
Sql = Sql & "RqCartouche.Ref_Plan_CLI = '" & UCase(MyReplace(Ref_Plan_CLI)) & "', "
Sql = Sql & "RqCartouche.[Ref_Piece_CLI] = '" & UCase(MyReplace(Ref_Piece_CLI)) & "', "
'Sql = Sql & "RqCartouche.Li = '" & MyReplace(txt10) & "', "
Sql = Sql & "RqCartouche.Client = '" & MyReplace(txt11) & "', "
Sql = Sql & "RqCartouche.Destinataire = '" & MyReplace(txt12) & "', "
Sql = Sql & "RqCartouche.Service ='" & MyReplace(txt13) & "', "
Sql = Sql & "RqCartouche.RefPF = '" & MyReplace(txt14) & "', "
Sql = Sql & "RqCartouche.RefP = '" & MyReplace(txt15) & "', "
Sql = Sql & "RqCartouche.DessineDate = " & MyReplaceDate(txt16) & ", "
Sql = Sql & "RqCartouche.DessineNOM ='" & MyReplace(txt17) & "', "
Sql = Sql & "RqCartouche.VerifieDate = " & MyReplaceDate(txt18) & ", "
Sql = Sql & "RqCartouche.VerifieNom = '" & MyReplace(txt19) & "', "
Sql = Sql & "RqCartouche.ApprouveDate = " & MyReplaceDate(txt20) & ", "
Sql = Sql & "RqCartouche.ApprouveNom ='" & MyReplace(txt21) & "', "
Sql = Sql & "RqCartouche.NbCartouche =" & txt22 & ", "
Sql = Sql & "RqCartouche.RefPieceClient= '" & MyReplace(txt23) & "', "
Sql = Sql & "RqCartouche.Masse= '" & MyReplace(TXT25) & "', "
Sql = Sql & "RqCartouche.BaseVehicule= '" & MyReplace(txt24) & "' "
Sql = Sql & "WHERE RqCartouche.T_indiceProjet.Id=" & Me.Tag & ";"
Con.Exequte Sql
IdIndiceProjet = Me.Tag
Sql = "SELECT T_indiceProjet.Id_Pieces FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"
Set Rs = Con.OpenRecordSet(Sql)
IdPieces = Rs!Id_Pieces
Sql = "SELECT T_Pieces.IdProjet FROM T_Pieces "
Sql = Sql & "WHERE T_Pieces.Id=" & IdPieces & ";"
Set Rs = Con.OpenRecordSet(Sql)
IdProjet = Rs!IdProjet

Sql = "DELETE T_Projet.* FROM T_Projet LEFT JOIN T_Pieces ON T_Projet.id = T_Pieces.IdProjet "
Sql = Sql & "WHERE T_Pieces.Id Is Null;"
Con.Exequte Sql
Execute = True
Sql = "SELECT T_indiceProjet.Pere FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Me.Tag & ";"
Set Rs = Con.OpenRecordSet(Sql)
If IdFils <> 0 Then Me.Tag = IdFils
IdFils = 0
If Rs!Pere > 0 Then
IdFils = Me.Tag
    Me.Tag = Rs!Pere
End If
 
  Sql = "SELECT T_Path.PathVar FROM T_Path WHERE T_Path.NameVar='PathOutils';"
    Set Rs = Con.OpenRecordSet(Sql)
    If Rs.EOF = False Then
        RepPlacheClous = "" & Rs!PathVar
    End If
Set Rs = Con.CloseRecordSet(Rs)
    

  
RepPlacheClous = RepPlacheClous & "\" & Me.PlanchClous
    
    Sql = "UPDATE T_indiceProjet SET T_indiceProjet.Cartouche = '" & MyReplace(RepPlacheClous) & "' "
    Sql = Sql & "WHERE T_indiceProjet.Id=" & Me.Tag & ";"
    Con.Exequte Sql
    
    
  Sql = "SELECT T_indiceProjet.Cartouche FROM T_indiceProjet WHERE T_indiceProjet.Id=" & Me.Tag & ";"
Set Rs = Con.OpenRecordSet(Sql)
RepPlacheClous = Rs!Cartouche

If boolAutoCAD = False Then
    MsgBox "Vos données ont bien été enregistrées, toute foie :" & vbCrLf & vbCrLf & "Vous ne possédez pas de licence AutoCad." & vbCrLf & "Vous ne pouvez pas reporter vos modifications" & vbCrLf & "sur vos différents plans. "
Else
bool_Plan_L_cartouches = True: bool_Plan_E_cartouches = True
 bool_Outil_L_cartouches = True: bool_Outil_E_cartouches = True
 ModifierUnCartouche Me.Tag
 bool_Plan_L_cartouches = False: bool_Plan_E_cartouches = False
 bool_Outil_L_cartouches = False: bool_Outil_E_cartouches = False
 End If
'Noquite = False
'
'Me.Hide
End Sub

Private Sub CommandButton8_Click()
Noquite = False
 Me.Hide
End Sub

Private Sub TextBox12_Change()

End Sub

Private Sub UserForm_Activate()
Execute = False
Noquite = True
End Sub

Private Sub UserForm_Initialize()
Dim Sql As String
Dim MyPath As String
Dim Rs As Recordset
Dim MyFichier As String
Set TableauPath = funPath
PlanchClous.Clear
MyPath = TableauPath.Item("PathOutils") & "\"
If Left(MyPath, 2) <> "\\" And Left(MyPath, 1) = "\" Then MyPath = TableauPath.Item("PathServer") & MyPath & "\"
If Right(MyPath, 2) = "\\" Then MyPath = Mid(MyPath, 1, Len(MyPath) - 1)



If Trim(MyPath) <> "" Then
MyFichier = Dir(MyPath & "*.dwg")
PlanchClous.AddItem ""
While MyFichier <> ""
PlanchClous.AddItem MyFichier
    MyFichier = Dir
 Wend
End If

boolCloseForm = True
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Cancel = Noquite
End Sub
