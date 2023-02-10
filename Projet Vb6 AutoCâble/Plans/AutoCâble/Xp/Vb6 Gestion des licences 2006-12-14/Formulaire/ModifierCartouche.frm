VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ModifierCartouche 
   Caption         =   "UserForm4"
   ClientHeight    =   11415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14475
   OleObjectBlob   =   "ModifierCartouche.dsx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ModifierCartouche"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim ChronoAnnee As String
Public IdProjet As Long
Public IdPieces As Long
Public IdIndiceProjet As Long
Public NbTxt As Long


Private Sub CommandButton1_Click()
Dim sql As String
Dim Rs As Recordset
CherchPices.Charge Me, "", True
Con.OpenConnetion db
sql = "SELECT RqCartouche.Projet AS txt1,  "
sql = sql & "RqCartouche.Vague AS txt2,  "
sql = sql & "RqCartouche.Equipement AS txt3,  "
sql = sql & "RqCartouche.Responsable AS txt4,  "
sql = sql & "RqCartouche.Ensemble AS txt5,  "
sql = sql & "RqCartouche.CleAc AS txt6,  "
sql = sql & "RqCartouche.PI AS txt7,  "
sql = sql & "RqCartouche.PL AS txt8,  "
sql = sql & "RqCartouche.[OU] AS txt9,  "
sql = sql & "RqCartouche.Li AS txt10,  "
sql = sql & "RqCartouche.Client AS txt11,  "
sql = sql & "RqCartouche.Destinataire AS txt12,  "
sql = sql & "RqCartouche.Service AS txt13,  "
sql = sql & "RqCartouche.RefPF AS txt14, "
sql = sql & " RqCartouche.RefP AS txt15,  "
sql = sql & "RqCartouche.DessineDate AS txt16,  "
sql = sql & "RqCartouche.DessineNOM AS txt17,  "
sql = sql & "RqCartouche.VerifieDate AS txt18,  "
sql = sql & "RqCartouche.VerifieNom AS txt19,  "
sql = sql & "RqCartouche.ApprouveDate AS txt20,  "
sql = sql & "RqCartouche.ApprouveNom AS txt21 "
sql = sql & "FROM RqCartouche "
sql = sql & "WHERE RqCartouche.T_indiceProjet.Id=" & Me.Tag & " ;"
Debug.Print sql
Set Rs = Con.OpenRecordSet(sql)
If Rs.EOF = False Then
For i = 1 To 21
    Me.Controls("txt" & CStr(i)) = "" & Rs.Fields("txt" & CStr(i))
Next i
End If

End Sub








Private Sub txt17_Change()
If Trim("" & Me.txt17.Text) = "" Then
     Me.txt16 = ""
Else
    Me.txt16 = Date
End If


End Sub



Private Sub txt6_Click()
Dim sql As String
Dim RsBaseNum As Recordset
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
Set RsBaseNum = ConNumPlan.OpenRecordSet(sql)

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
Set RsBaseNum = ConNumPlan.OpenRecordSet(sql)

While RsBaseNum.EOF = False
    txt8.AddItem "" & RsBaseNum![Clé ty] & "_" & RsBaseNum![Clé ac] & "_" & RsBaseNum![Année] & _
    "_" & RsBaseNum![Clé Ch] & "_" & RsBaseNum![rv] & "_" & RsBaseNum![Rév]
    txt8.List(txt8.ListCount - 1, 1) = "" & RsBaseNum![Clé Ch]
    txt8.List(txt8.ListCount - 1, 1) = "" & RsBaseNum![Clé Ch]
    txt8.List(txt8.ListCount - 1, 2) = "" & RsBaseNum![Red_Nom] & " " & RsBaseNum![Red_P_Nom]
    txt8.List(txt8.ListCount - 1, 3) = "" & RsBaseNum![Verif_Nom] & " " & RsBaseNum![Verif_P_Nom]
    txt8.List(txt8.ListCount - 1, 3) = "" & RsBaseNum![Apr_Nom] & " " & RsBaseNum![Apr_P_Nom]
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
Set RsBaseNum = ConNumPlan.OpenRecordSet(sql)

While RsBaseNum.EOF = False
    txt9.AddItem "" & RsBaseNum![Clé ty] & "_" & RsBaseNum![Clé ac] & "_" & RsBaseNum![Année] & _
    "_" & RsBaseNum![Clé Ch] & "_" & RsBaseNum![rv] & "_" & RsBaseNum![Rév]
    txt9.List(txt9.ListCount - 1, 1) = "" & RsBaseNum![Clé Ch]
    txt9.List(txt9.ListCount - 1, 2) = "" & RsBaseNum![Red_Nom] & " " & RsBaseNum![Red_P_Nom]
    txt9.List(txt9.ListCount - 1, 3) = "" & RsBaseNum![Verif_Nom] & " " & RsBaseNum![Verif_P_Nom]
    txt9.List(txt9.ListCount - 1, 3) = "" & RsBaseNum![Apr_Nom] & " " & RsBaseNum![Apr_P_Nom]
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
Set RsBaseNum = ConNumPlan.OpenRecordSet(sql)

While RsBaseNum.EOF = False
    txt10.AddItem "" & RsBaseNum![Clé ty] & "_" & RsBaseNum![Clé ac] & "_" & RsBaseNum![Année] & _
    "_" & RsBaseNum![Clé Ch] & "_" & RsBaseNum![rv] & "_" & RsBaseNum![Rév]
    txt10.List(txt10.ListCount - 1, 1) = "" & RsBaseNum![Clé Ch]
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
UserForm3.Show vbModal
Maj Me.txt1

End Sub

Private Sub CommandButton7_Click()
Dim sql As String
Dim Rs As Recordset
NbTxt = 21
If ValideChampsTexte(Me, NbTxt) = False Then Exit Sub

  
sql = "INSERT INTO T_indiceProjet ( [Indice],Id_Pieces,PI,PL,OU,  Li,  IdStatus,  "
sql = sql & "Client, Destinataire, Service, DessineDate, DessineNOM, VerifieDate,  "
sql = sql & "VerifieNom, ApprouveDate, ApprouveNom, Responsable, Vague, Equipement,  "
sql = sql & "Ensemble, CleAc,RefPF,RefP ) "
sql = sql & "values(  "
'sql = sql & "'" & MyReplace(txt17) & "',  "
sql = sql & "'" & MyReplace(txt7.List(txt7.ListIndex, 5)) & "' "
sql = sql & "," & IdPieces & ", "
sql = sql & "'" & MyReplace(txt7) & "',  "
sql = sql & "'" & MyReplace(txt8) & "',  "
sql = sql & "'" & MyReplace(txt9) & "',  "
sql = sql & "'" & MyReplace(txt10) & "', "
sql = sql & "1 ,  "
sql = sql & "'" & MyReplace(txt11.Text) & "',  "
sql = sql & "'" & MyReplace(txt12) & "',  "
sql = sql & "'" & MyReplace(txt13) & "',  "
sql = sql & MyReplaceDate(txt16) & ",  "
sql = sql & "'" & MyReplace(txt17) & "',  "
sql = sql & MyReplaceDate(txt18) & ",  "
sql = sql & "'" & MyReplace((txt19)) & "',  "
sql = sql & MyReplaceDate(txt20) & ",  "
sql = sql & "'" & MyReplace(txt21) & "',  "
sql = sql & "'" & MyReplace(txt4) & "',  "
sql = sql & "'" & MyReplace(txt2) & "',  "
sql = sql & "'" & MyReplace(txt3) & "',  "
'Sql = Sql & "'" & MyReplace(txt19) & "',  "
sql = sql & "'" & MyReplace(txt5) & "',  "
sql = sql & "'" & MyReplace(txt6) & "',  "
sql = sql & "'" & MyReplace(txt14) & "',  "
sql = sql & "'" & MyReplace(txt15) & "'); "
Con.Exequte sql


ImportXls.Charge Me



End Sub

Private Sub UserForm_Activate()
Dim Rs As Recordset
Dim RsBaseNum As Recordset
Dim sql As String
Dim indexClient As Long
Dim RqChronoAnnee As String
RqChronoAnne = "[Chrono Requête " & Format(Date, "yyyy]")
ChronoAnnee = "[Chrono" & Format(Date, "yyyy]")
sql = "SELECT T_Clients.Client, T_Clients.Formulaire "
sql = sql & "FROM T_Clients "
sql = sql & "ORDER BY T_Clients.Client;"
Con.OpenConnetion db
ConNumPlan.OpenConnetion DbNumPlan
Set Rs = Con.OpenRecordSet(sql)
 Me.txt11.AddItem ""
While Rs.EOF = False
    Me.txt11.AddItem Trim("" & Rs!Client)
        If Me.txt11.ListCount = 1 Then Me.txt11.Text = Trim("" & Rs!Client)

    Rs.MoveNext
Wend




'Sql = "SELECT " & ChronoAnnee & ".[Clé ty], " & ChronoAnnee & ".[Clé ac],  "
'Sql = Sql & ChronoAnnee & ".Année, " & ChronoAnnee & ".[Clé Ch], "
'Sql = Sql & " " & ChronoAnnee & ".rv, " & ChronoAnnee & ".Rév "
'Sql = Sql & "FROM " & ChronoAnnee & " "
'Sql = Sql & "WHERE " & ChronoAnnee & ".[Clé ty] = 'LI' "
'Sql = Sql & "ORDER BY " & ChronoAnnee & ".[Clé Ch] DESC;"
'Set RsBaseNum = ConNumPlan.OpenRecordSet(Sql)
'
'While RsBaseNum.EOF = False
'    ComboBox1.AddItem "" & RsBaseNum![Clé Ch]
'    RsBaseNum.MoveNext
'Wend

sql = "SELECT Activité.[Clé ac], Responsables.Nom, Responsables.Prénom, " & ChronoAnnee & ".Destinataire "
sql = sql & "FROM (Responsables RIGHT JOIN Activité ON Responsables.[Clé res] = Activité.[Clé re]) INNER JOIN " & ChronoAnnee & " ON Activité.[Clé ac] = " & ChronoAnnee & ".[Clé ac] "
sql = sql & "WHERE " & ChronoAnnee & ".[Clé ty]='PL' "
sql = sql & "OR " & ChronoAnnee & ".[Clé ty]='OU' "
sql = sql & "OR " & ChronoAnnee & ".[Clé ty]='LI' "
sql = sql & "OR " & ChronoAnnee & ".[Clé ty]='PI' "
sql = sql & "GROUP BY Activité.[Clé ac], Responsables.Nom, Responsables.Prénom, " & ChronoAnnee & ".Destinataire "
sql = sql & "ORDER BY Activité.[Clé ac] DESC;"

sql = "SELECT Activité.[Clé ac], Activité.[Date ac], Activité.[Clé re],  "
sql = sql & "Activité.Client, Activité.Int, Activité.[Lib ac], Activité.[Obs ac],  "
sql = sql & "Activité.[Clé tyac], Activité.[St ac], Activité.[Clé pr],  "
sql = sql & "Activité.[Clé ca], Activité.[vid ac]  "
sql = sql & "FROM Activité INNER JOIN " & ChronoAnnee & " ON Activité.[Clé ac] = " & ChronoAnnee & ".[Clé ac] "
sql = sql & "WHERE " & ChronoAnnee & ".[Clé ty]='PL'  "
sql = sql & "OR " & ChronoAnnee & ".[Clé ty]='OU'  "
sql = sql & "OR " & ChronoAnnee & ".[Clé ty]='LI'  "
sql = sql & "OR " & ChronoAnnee & ".[Clé ty]='PI' "
sql = sql & "ORDER BY Activité.[Clé ac] DESC;"


sql = "SELECT Activité.[Clé ac], Responsables.Nom, Responsables.Prénom, Activité.Int "
sql = sql & "FROM Responsables INNER JOIN (Activité INNER JOIN " & ChronoAnnee & "  "
sql = sql & "ON Activité.[Clé ac] = " & ChronoAnnee & ".[Clé ac])  "
sql = sql & "ON Responsables.[Clé res] = Activité.[Clé re] "
sql = sql & "WHERE " & ChronoAnnee & ".[Clé ty]='PL'  "
sql = sql & "Or " & ChronoAnnee & ".[Clé ty]='OU'  "
sql = sql & "Or " & ChronoAnnee & ".[Clé ty]='LI'  "
sql = sql & "Or " & ChronoAnnee & ".[Clé ty]='PI' "
sql = sql & "GROUP BY Activité.[Clé ac], Responsables.Nom, Responsables.Prénom, Activité.Int "
sql = sql & "ORDER BY Activité.[Clé ac] DESC;"



''
'SELECT " & ChronoAnnee & ".[Clé ac], " & ChronoAnnee & ".[Clé ty]
'FROM " & ChronoAnnee & "
'GROUP BY " & ChronoAnnee & ".[Clé ac], " & ChronoAnnee & ".[Clé ty]
''HAVING (((" & ChronoAnnee & ".[Clé ty]) = "PL")) Or (((" & ChronoAnnee & ".[Clé ty]) = "OU")) Or (((" & ChronoAnnee & ".[Clé ty]) = "LI")) Or (((" & ChronoAnnee & ".[Clé ty]) = "PI"))
'ORDER BY " & ChronoAnnee & ".[Clé ac] DESC;



Set RsBaseNum = ConNumPlan.OpenRecordSet(sql)


While RsBaseNum.EOF = False
    txt6.AddItem CStr(RsBaseNum![Clé ac])
    txt6.List(txt6.ListCount - 1, 1) = Trim("" & RsBaseNum!NOM) & " " & Trim("" & RsBaseNum!Prénom)
      txt6.List(txt6.ListCount - 1, 2) = Trim("" & RsBaseNum!Int)
 
    RsBaseNum.MoveNext
Wend
Set Rs = Con.CloseRecordSet(Rs)
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Con.CloseConnection
ConNumPlan.CloseConnection
End Sub


