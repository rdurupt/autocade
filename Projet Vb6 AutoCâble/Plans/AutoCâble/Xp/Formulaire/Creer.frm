VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Creer 
   Caption         =   "Créer un nouveau plan de câblage:"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11670
   OleObjectBlob   =   "Creer.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Creer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ChronoAnnee As String
Public IdProjet As Long
Public IdPieces As Long
Public IdIndiceProjet As Long
Public NbTxt As Long
Dim Noquite As Boolean



Private Sub CommandButton10_Click()
Dim sql As String
Dim Rs As Recordset
CherchPices.Charge Me, " LiAutoCadSave <>  Null and NbErr=0 AND Pere=0 and CleAc=" & Val(txt6), True
Unload CherchPices

sql = "SELECT T_indiceProjet.Id, [PI] & '_' & [PI_Indice] AS Piece FROM T_indiceProjet "
sql = sql & "WHERE T_indiceProjet.Id=" & Val(Me.Tag) & ";"
Set Rs = Con.OpenRecordSet(sql)
If Rs.EOF = False Then
Me.Pere = Rs!Piece
End If
Me.Pere.Tag = Val(Me.Tag)
Set Rs = Con.CloseRecordSet(Rs)

End Sub

Private Sub CommandButton8_Click()
Noquite = False
Me.Hide

End Sub

Private Sub tx15_Change()

End Sub

Private Sub tx19_Change()

End Sub

Private Sub tx20_Change()

End Sub






Private Sub Label35_Click()

End Sub

Private Sub txt17_Change()
If Trim("" & Me.txt17.Text) = "" Then
     Me.txt16 = ""
Else
    Me.txt16 = Format(Date, "dd/mm/yyyy")
End If


End Sub



Private Sub txt6_Click()
Dim sql As String
Dim RsBaseNum As Recordset
ConBaseNum.OpenConnetion DbNumPlan

txt7.Clear
txt8.Clear
txt9.Clear
txt10.Clear

sql = "SELECT " & ChronoAnnee & ".[Clé ty], " & ChronoAnnee & ".[Clé ac],  "
sql = sql & ChronoAnnee & ".Année, " & ChronoAnnee & ".[Clé Ch], "
sql = sql & " " & ChronoAnnee & ".rv, " & ChronoAnnee & ".Rév "
sql = sql & "FROM " & ChronoAnnee & "  "

sql = "SELECT Agent_2.[Nom ag] AS Red_Nom,  "
sql = sql & "Agent_2.[Prénom ag] AS Red_P_Nom,  "
sql = sql & "Agent_1.[Nom ag] AS Verif_Nom,  "
sql = sql & "Agent_1.[Prénom ag] AS Verif_P_Nom,  "
sql = sql & "Agent.[Nom ag] AS Apr_Nom,  "
sql = sql & "Agent.[Prénom ag] AS Apr_P_Nom,  "
sql = sql & "" & ChronoAnnee & ".[Clé ty], " & ChronoAnnee & ".[Clé ac],  "
sql = sql & "" & ChronoAnnee & ".Année, " & ChronoAnnee & ".[Clé Ch],  "
sql = sql & "" & ChronoAnnee & ".rv, " & ChronoAnnee & ".Rév  "
sql = sql & "FROM ((" & ChronoAnnee & " INNER JOIN Agent ON " & ChronoAnnee & ".[Clé ap] = Agent.[Clé ag])  "
sql = sql & "INNER JOIN Agent AS Agent_1 ON " & ChronoAnnee & ".[Clé ve] = Agent_1.[Clé ag])  "
sql = sql & "INNER JOIN Agent AS Agent_2 ON " & ChronoAnnee & ".[Clé re] = Agent_2.[Clé ag] "
sql = sql & "WHERE " & ChronoAnnee & ".[Clé ac] = " & txt6.Text & " and " & ChronoAnnee & ".[Clé ty] = 'PI' "
sql = sql & "ORDER BY " & ChronoAnnee & ".[Clé Ch] DESC;"
Set RsBaseNum = ConBaseNum.OpenRecordSet(sql)

While RsBaseNum.EOF = False

    txt7.AddItem "" & RsBaseNum![Clé ty] & "_" & RsBaseNum![Clé ac] & "_" & RsBaseNum![Année] & _
    "_" & RsBaseNum![Clé Ch] & "_" & RsBaseNum![rv] & "_" & RsBaseNum![Rév]
    txt7.List(txt7.ListCount - 1, 1) = "" & RsBaseNum![Clé Ch]
    If Len(Trim(" " & RsBaseNum![Red_P_Nom])) > 0 Then
        txt7.List(txt7.ListCount - 1, 2) = "" & RsBaseNum![Red_Nom] & " " & Left(UCase(RsBaseNum![Red_P_Nom]), 1)
    Else
         txt7.List(txt7.ListCount - 1, 2) = "" & RsBaseNum![Red_Nom]
    End If
    If Len(Trim(" " & RsBaseNum![Verif_P_Nom])) > 0 Then
         txt7.List(txt7.ListCount - 1, 3) = "" & RsBaseNum![Verif_Nom] & " " & Left(UCase(RsBaseNum![Verif_P_Nom]), 1)
    Else
         txt7.List(txt7.ListCount - 1, 3) = "" & RsBaseNum![Verif_Nom]
    End If
   
   If Len(Trim(" " & RsBaseNum![Apr_P_Nom])) > 0 Then
         txt7.List(txt7.ListCount - 1, 4) = "" & RsBaseNum![Apr_Nom] & " " & Left(UCase(RsBaseNum![Apr_P_Nom]), 1)
    Else
        txt7.List(txt7.ListCount - 1, 4) = "" & RsBaseNum![Apr_Nom]
    End If
    txt7.List(txt7.ListCount - 1, 5) = "" & RsBaseNum![Rév]

    RsBaseNum.MoveNext
Wend

sql = "SELECT " & ChronoAnnee & ".[Clé ty], " & ChronoAnnee & ".[Clé ac],  "
sql = sql & ChronoAnnee & ".Année, " & ChronoAnnee & ".[Clé Ch], "
sql = sql & " " & ChronoAnnee & ".rv, " & ChronoAnnee & ".Rév "
sql = sql & "FROM " & ChronoAnnee & " "

sql = "SELECT Agent_2.[Nom ag] AS Red_Nom,  "
sql = sql & "Agent_2.[Prénom ag] AS Red_P_Nom,  "
sql = sql & "Agent_1.[Nom ag] AS Verif_Nom,  "
sql = sql & "Agent_1.[Prénom ag] AS Verif_P_Nom,  "
sql = sql & "Agent.[Nom ag] AS Apr_Nom,  "
sql = sql & "Agent.[Prénom ag] AS Apr_P_Nom,  "
sql = sql & "" & ChronoAnnee & ".[Clé ty], " & ChronoAnnee & ".[Clé ac],  "
sql = sql & "" & ChronoAnnee & ".Année, " & ChronoAnnee & ".[Clé Ch],  "
sql = sql & "" & ChronoAnnee & ".rv, " & ChronoAnnee & ".Rév  "
sql = sql & "FROM ((" & ChronoAnnee & " INNER JOIN Agent ON " & ChronoAnnee & ".[Clé ap] = Agent.[Clé ag])  "
sql = sql & "INNER JOIN Agent AS Agent_1 ON " & ChronoAnnee & ".[Clé ve] = Agent_1.[Clé ag])  "
sql = sql & "INNER JOIN Agent AS Agent_2 ON " & ChronoAnnee & ".[Clé re] = Agent_2.[Clé ag] "
sql = sql & "WHERE " & ChronoAnnee & ".[Clé ac] = " & txt6.Text & " and " & ChronoAnnee & ".[Clé ty] = 'PL' "
sql = sql & "ORDER BY " & ChronoAnnee & ".[Clé Ch] DESC;"
Set RsBaseNum = ConBaseNum.OpenRecordSet(sql)

While RsBaseNum.EOF = False
    txt8.AddItem "" & RsBaseNum![Clé ty] & "_" & RsBaseNum![Clé ac] & "_" & RsBaseNum![Année] & _
    "_" & RsBaseNum![Clé Ch] & "_" & RsBaseNum![rv] & "_" & RsBaseNum![Rév]
    txt8.List(txt8.ListCount - 1, 1) = "" & RsBaseNum![Clé Ch]
    txt8.List(txt8.ListCount - 1, 1) = "" & RsBaseNum![Clé Ch]
    txt8.List(txt8.ListCount - 1, 2) = "" & RsBaseNum![Red_Nom] & " " & RsBaseNum![Red_P_Nom]
    txt8.List(txt8.ListCount - 1, 3) = "" & RsBaseNum![Verif_Nom] & " " & RsBaseNum![Verif_P_Nom]
    txt8.List(txt8.ListCount - 1, 4) = "" & RsBaseNum![Apr_Nom] & " " & RsBaseNum![Apr_P_Nom]
    txt8.List(txt8.ListCount - 1, 5) = "" & RsBaseNum![Rév]
    RsBaseNum.MoveNext
Wend



sql = "SELECT " & ChronoAnnee & ".[Clé ty], " & ChronoAnnee & ".[Clé ac],  "
sql = sql & ChronoAnnee & ".Année, " & ChronoAnnee & ".[Clé Ch], "
sql = sql & " " & ChronoAnnee & ".rv, " & ChronoAnnee & ".Rév "
sql = sql & "FROM " & ChronoAnnee & " "

sql = "SELECT Agent_2.[Nom ag] AS Red_Nom,  "
sql = sql & "Agent_2.[Prénom ag] AS Red_P_Nom,  "
sql = sql & "Agent_1.[Nom ag] AS Verif_Nom,  "
sql = sql & "Agent_1.[Prénom ag] AS Verif_P_Nom,  "
sql = sql & "Agent.[Nom ag] AS Apr_Nom,  "
sql = sql & "Agent.[Prénom ag] AS Apr_P_Nom,  "
sql = sql & "" & ChronoAnnee & ".[Clé ty], " & ChronoAnnee & ".[Clé ac],  "
sql = sql & "" & ChronoAnnee & ".Année, " & ChronoAnnee & ".[Clé Ch],  "
sql = sql & "" & ChronoAnnee & ".rv, " & ChronoAnnee & ".Rév  "
sql = sql & "FROM ((" & ChronoAnnee & " INNER JOIN Agent ON " & ChronoAnnee & ".[Clé ap] = Agent.[Clé ag])  "
sql = sql & "INNER JOIN Agent AS Agent_1 ON " & ChronoAnnee & ".[Clé ve] = Agent_1.[Clé ag])  "
sql = sql & "INNER JOIN Agent AS Agent_2 ON " & ChronoAnnee & ".[Clé re] = Agent_2.[Clé ag] "
sql = sql & "WHERE " & ChronoAnnee & ".[Clé ac] = " & txt6.Text & " and " & ChronoAnnee & ".[Clé ty] = 'OU' "
sql = sql & "ORDER BY " & ChronoAnnee & ".[Clé Ch] DESC;"
Set RsBaseNum = ConBaseNum.OpenRecordSet(sql)

While RsBaseNum.EOF = False
    txt9.AddItem "" & RsBaseNum![Clé ty] & "_" & RsBaseNum![Clé ac] & "_" & RsBaseNum![Année] & _
    "_" & RsBaseNum![Clé Ch] & "_" & RsBaseNum![rv] & "_" & RsBaseNum![Rév]
    txt9.List(txt9.ListCount - 1, 1) = "" & RsBaseNum![Clé Ch]
    txt9.List(txt9.ListCount - 1, 2) = "" & RsBaseNum![Red_Nom] & " " & RsBaseNum![Red_P_Nom]
    txt9.List(txt9.ListCount - 1, 3) = "" & RsBaseNum![Verif_Nom] & " " & RsBaseNum![Verif_P_Nom]
    txt9.List(txt9.ListCount - 1, 4) = "" & RsBaseNum![Apr_Nom] & " " & RsBaseNum![Apr_P_Nom]
    txt9.List(txt9.ListCount - 1, 5) = " " & RsBaseNum![Rév]
    RsBaseNum.MoveNext
Wend
sql = "SELECT " & ChronoAnnee & ".[Clé ty], " & ChronoAnnee & ".[Clé ac],  "
sql = sql & ChronoAnnee & ".Année, " & ChronoAnnee & ".[Clé Ch], "
sql = sql & " " & ChronoAnnee & ".rv, " & ChronoAnnee & ".Rév "
sql = sql & "FROM " & ChronoAnnee & " "

sql = "SELECT Agent_2.[Nom ag] AS Red_Nom,  "
sql = sql & "Agent_2.[Prénom ag] AS Red_P_Nom,  "
sql = sql & "Agent_1.[Nom ag] AS Verif_Nom,  "
sql = sql & "Agent_1.[Prénom ag] AS Verif_P_Nom,  "
sql = sql & "Agent.[Nom ag] AS Apr_Nom,  "
sql = sql & "Agent.[Prénom ag] AS Apr_P_Nom,  "
sql = sql & "" & ChronoAnnee & ".[Clé ty], " & ChronoAnnee & ".[Clé ac],  "
sql = sql & "" & ChronoAnnee & ".Année, " & ChronoAnnee & ".[Clé Ch],  "
sql = sql & "" & ChronoAnnee & ".rv, " & ChronoAnnee & ".Rév  "
sql = sql & "FROM ((" & ChronoAnnee & " INNER JOIN Agent ON " & ChronoAnnee & ".[Clé ap] = Agent.[Clé ag])  "
sql = sql & "INNER JOIN Agent AS Agent_1 ON " & ChronoAnnee & ".[Clé ve] = Agent_1.[Clé ag])  "
sql = sql & "INNER JOIN Agent AS Agent_2 ON " & ChronoAnnee & ".[Clé re] = Agent_2.[Clé ag] "
sql = sql & "WHERE " & ChronoAnnee & ".[Clé ac] = " & txt6.Text & " and " & ChronoAnnee & ".[Clé ty] = 'LI' "
sql = sql & "ORDER BY " & ChronoAnnee & ".[Clé Ch] DESC;"
Set RsBaseNum = ConBaseNum.OpenRecordSet(sql)

While RsBaseNum.EOF = False
    txt10.AddItem "" & RsBaseNum![Clé ty] & "_" & RsBaseNum![Clé ac] & "_" & RsBaseNum![Année] & _
    "_" & RsBaseNum![Clé Ch] & "_" & RsBaseNum![rv] & "_" & RsBaseNum![Rév]
    txt10.List(txt10.ListCount - 1, 1) = "" & RsBaseNum![Clé Ch]
     txt10.List(txt10.ListCount - 1, 2) = " " & RsBaseNum![Rév]
    RsBaseNum.MoveNext
Wend

If Me.txt6.ListIndex <> -1 Then
    Me.txt12 = txt6.List(Me.txt6.ListIndex, 2)
Else
     Me.txt12 = ""
End If

If Me.txt6.ListIndex <> -1 Then
    Me.txt4 = txt6.List(Me.txt6.ListIndex, 1)
Else
     Me.txt4 = ""
End If
ConBaseNum.CloseConnection
End Sub






Private Sub tx11_Change()

End Sub

Private Sub tx11_Click()

End Sub

Private Sub txt7_Change()
If txt7.ListIndex > -1 Then
    txt17 = txt7.List(txt7.ListIndex, 2)
   txt19 = txt7.List(txt7.ListIndex, 3)
    txt21 = txt7.List(txt7.ListIndex, 4)
Else
    txt17 = ""
    txt19 = ""
    txt21 = ""
End If
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

Private Sub CommandButton6_Click()
UserForm3.Show
Unload UserForm3
If Me.txt1 <> "" Then Maj Me.txt1

End Sub

Private Sub CommandButton7_Click()
Dim sql As String
Dim Rs As Recordset
Dim pose As Long
Dim txt As String

NbTxt = 21
If ValideChampsTexte(Me, NbTxt) = False Then Exit Sub

sql = "SELECT T_Projet.id FROM T_Projet "


sql = sql & "WHERE T_Projet.Projet='" & MyReplace(txt1) & "';"
Set Rs = Con.OpenRecordSet(sql)
If Rs.EOF = True Then

    sql = "INSERT INTO T_Projet ( Projet ) "
    sql = sql & "VALUES( '" & MyReplace(txt1) & "');"
    Con.Exequte sql
End If
Rs.Requery

IdProjet = Rs!Id

pose = InStr(1, txt7, "_Rév.:_")
If pose = 0 Then
    txt = txt7
Else
    txt = Mid(txt7, 1, pose - 1)
End If
sql = "SELECT T_Pieces.Id, T_Projet.Projet "
sql = sql & "FROM T_Projet INNER JOIN T_Pieces ON T_Projet.id = T_Pieces.IdProjet "
sql = sql & "WHERE T_Pieces.Description='" & MyReplace(txt) & "';"
Set Rs = Con.OpenRecordSet(sql)
If Rs.EOF = False Then
    MsgBox "La pièce : " & txt & " existe déjà dans le projet : " & Rs!Projet & vbCrLf & "Opération d'ajout annulée"
    Exit Sub
End If
sql = "INSERT INTO T_Pieces ( IdProjet, Description )"
sql = sql & "VALUES( " & IdProjet & ", '" & MyReplace(txt) & "');"
 Con.Exequte sql
 Rs.Requery
 IdPieces = Rs!Id
 sql = "INSERT INTO T_indiceProjet ( Id_Pieces, Vague, Equipement, Responsable, Ensemble, "
 sql = sql & "CleAc, PI, PL,PL_Indice, OU,OU_Indice, Li,LI_Indice, Client, Destinataire, Service, RefPF, RefP, DessineDate,  "
 sql = sql & "DessineNOM, VerifieDate, VerifieNom, ApprouveDate, ApprouveNom ,PI_Indice "
 
 If Trim("" & Me.Pere) <> "" Then
    sql = sql & ",pere "
 End If
 
 sql = sql & ") "
 
 sql = sql & "VALUES ( " & IdPieces & " , '" & MyReplace(txt2) & "' , '" & MyReplace(txt3) & "', "
 sql = sql & "'" & MyReplace(txt4) & "' , '" & MyReplace(txt5) & "' , '" & MyReplace(txt6) & "' ,  "
 sql = sql & "'" & MyReplace(txt) & "' , "
 pose = InStr(1, txt8, "_Rév.:_")
If pose = 0 Then
    txt = txt8
Else
    txt = Mid(txt8, 1, pose - 1)
End If
sql = sql & "'" & MyReplace(txt) & "' ,"
sql = sql & "'" & MyReplace(txt8.List(txt8.ListIndex, 5)) & "',"

 pose = InStr(1, txt9, "_Rév.:_")
If pose = 0 Then
    txt = txt9
Else
    txt = Mid(txt9, 1, pose - 1)
End If
 sql = sql & "'" & MyReplace(txt) & "' ,  "
 sql = sql & "'" & MyReplace(txt9.List(txt9.ListIndex, 5)) & "',"

 pose = InStr(1, txt10, "_Rév.:_")
If pose = 0 Then
    txt = txt10
Else
    txt = Mid(txt10, 1, pose - 1)
End If
sql = sql & "'" & MyReplace(txt) & "', "
sql = sql & "'" & MyReplace(txt10.List(txt10.ListIndex, 2)) & "',"


 sql = sql & "'" & MyReplace(txt11) & "' , '" & MyReplace(txt2) & "' ,  "
 sql = sql & "'" & MyReplace(txt13) & "' , '" & MyReplace(txt14) & "' , '" & MyReplace(txt15) & "' ,  "
 sql = sql & "" & MyReplaceDate(txt16) & " , '" & MyReplace(txt17) & "' , " & MyReplaceDate(txt18) & " ,  "
 sql = sql & "'" & MyReplace(txt19) & "' , " & MyReplaceDate(txt20) & " , '" & MyReplace(txt21) & "','" & txt7.List(txt7.ListIndex, 5) & "' "
 
 If Trim("" & Me.Pere) <> "" Then
    sql = sql & "," & Me.Pere.Tag
 End If
 
 sql = sql & ");"
Con.Exequte sql
sql = "SELECT T_indiceProjet.Id "
sql = sql & "FROM T_indiceProjet "
sql = sql & "WHERE T_indiceProjet.Id_Pieces=" & IdPieces & " ;"
Set Rs = Con.OpenRecordSet(sql)
IdIndiceProjet = Rs!Id
Me.Hide
If Trim("" & Me.Pere) <> "" Then
Modifier.Charge Me
Unload Modifier
Else
    ImportXls.Charge Me
    Unload ImportXls
End If
End Sub

Private Sub UserForm_Activate()
Dim Rs As Recordset
Dim RsBaseNum As Recordset
Dim sql As String
Dim indexClient As Long
Dim RqChronoAnnee As String
ConBaseNum.OpenConnetion DbNumPlan
Noquite = True
RqChronoAnne = "[Chrono Requête " & Format(Date, "yyyy]")
ChronoAnnee = "Chrono" & Format(Date, "yyyy")
sql = "SELECT T_Clients.Client, T_Clients.Formulaire "
sql = sql & "FROM T_Clients "
sql = sql & "ORDER BY T_Clients.Client;"


Set Rs = Con.OpenRecordSet(sql)
 Me.txt11.AddItem ""
While Rs.EOF = False
    Me.txt11.AddItem Trim("" & Rs!Client)
        If Me.txt11.ListCount = 1 Then Me.txt11.Text = Trim("" & Rs!Client)

    Rs.MoveNext
Wend






sql = "SELECT Activité.[Clé ac], Responsables.Nom, Responsables.Prénom, Activité.Int "
sql = sql & "FROM Responsables INNER JOIN (Activité INNER JOIN " & ChronoAnnee & "  "
sql = sql & "ON Activité.[Clé ac] = " & ChronoAnnee & ".[Clé ac])  "
sql = sql & "ON Responsables.[Clé res] = Activité.[Clé re] "
sql = sql & "WHERE " & ChronoAnnee & ".[Clé ty]='PL'  "
sql = sql & "Or " & ChronoAnnee & ".[Clé ty]='OU'  "
sql = sql & "or " & ChronoAnnee & ".[Clé ty]='LI'  "
sql = sql & "Or " & ChronoAnnee & ".[Clé ty]='PI' "
sql = sql & "GROUP BY Activité.[Clé ac], Responsables.Nom, Responsables.Prénom, Activité.Int "
sql = sql & "ORDER BY Activité.[Clé ac] DESC;"


 

Set RsBaseNum = ConBaseNum.OpenRecordSet(sql)


While RsBaseNum.EOF = False
    txt6.AddItem CStr(RsBaseNum![Clé ac])
    txt6.List(txt6.ListCount - 1, 1) = Trim("" & RsBaseNum!NOM) & " " & Trim("" & RsBaseNum!Prénom)
      txt6.List(txt6.ListCount - 1, 2) = Trim("" & RsBaseNum!Int)
 
    RsBaseNum.MoveNext
Wend
Set Rs = Con.CloseRecordSet(Rs)
ConBaseNum.CloseConnection
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Cancel = Noquite
End Sub
