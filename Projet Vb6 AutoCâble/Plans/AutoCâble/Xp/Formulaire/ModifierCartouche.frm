VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ModifierCartouche 
   Caption         =   "UserForm4"
   ClientHeight    =   11415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14475
   OleObjectBlob   =   "ModifierCartouche.frx":0000
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
Dim Sql As String
Dim Rs As Recordset
CherchPices.Charge Me, "", True
Con.OpenConnetion db
Sql = "SELECT RqCartouche.Projet AS txt1,  "
Sql = Sql & "RqCartouche.Vague AS txt2,  "
Sql = Sql & "RqCartouche.Equipement AS txt3,  "
Sql = Sql & "RqCartouche.Responsable AS txt4,  "
Sql = Sql & "RqCartouche.Ensemble AS txt5,  "
Sql = Sql & "RqCartouche.CleAc AS txt6,  "
Sql = Sql & "RqCartouche.PI AS txt7,  "
Sql = Sql & "RqCartouche.PL AS txt8,  "
Sql = Sql & "RqCartouche.[OU] AS txt9,  "
Sql = Sql & "RqCartouche.Li AS txt10,  "
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
Sql = Sql & "RqCartouche.ApprouveNom AS txt21 "
Sql = Sql & "FROM RqCartouche "
Sql = Sql & "WHERE RqCartouche.T_indiceProjet.Id=" & Me.Tag & " ;"
Debug.Print Sql
Set Rs = Con.OpenRecordSet(Sql)
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
Dim Sql As String
Dim RsBaseNum As Recordset
txt7.Clear
txt8.Clear
txt9.Clear
txt10.Clear

Sql = "SELECT " & ChronoAnnee & ".[Clé ty], " & ChronoAnnee & ".[Clé ac],  "
Sql = Sql & ChronoAnnee & ".Année, " & ChronoAnnee & ".[Clé Ch], "
Sql = Sql & " " & ChronoAnnee & ".rv, " & ChronoAnnee & ".Rév "
Sql = Sql & "FROM " & ChronoAnnee & "  "

Sql = "SELECT Agent_2.[Nom ag] AS Red_Nom,  "
Sql = Sql & "Agent_2.[Prénom ag] AS Red_P_Nom,  "
Sql = Sql & "Agent_1.[Nom ag] AS Verif_Nom,  "
Sql = Sql & "Agent_1.[Prénom ag] AS Verif_P_Nom,  "
Sql = Sql & "Agent.[Nom ag] AS Apr_Nom,  "
Sql = Sql & "Agent.[Prénom ag] AS Apr_P_Nom,  "
Sql = Sql & "" & ChronoAnnee & ".[Clé ty], " & ChronoAnnee & ".[Clé ac],  "
Sql = Sql & "" & ChronoAnnee & ".Année, " & ChronoAnnee & ".[Clé Ch],  "
Sql = Sql & "" & ChronoAnnee & ".rv, " & ChronoAnnee & ".Rév  "
Sql = Sql & "FROM ((" & ChronoAnnee & " INNER JOIN Agent ON " & ChronoAnnee & ".[Clé ap] = Agent.[Clé ag])  "
Sql = Sql & "INNER JOIN Agent AS Agent_1 ON " & ChronoAnnee & ".[Clé ve] = Agent_1.[Clé ag])  "
Sql = Sql & "INNER JOIN Agent AS Agent_2 ON " & ChronoAnnee & ".[Clé re] = Agent_2.[Clé ag] "

Sql = Sql & "WHERE " & ChronoAnnee & ".[Clé ac] = " & txt6.Text & " and " & ChronoAnnee & ".[Clé ty] = 'PI' "
Sql = Sql & "ORDER BY " & ChronoAnnee & ".[Clé Ch] DESC;"
Set RsBaseNum = ConNumPlan.OpenRecordSet(Sql)

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

Sql = "SELECT " & ChronoAnnee & ".[Clé ty], " & ChronoAnnee & ".[Clé ac],  "
Sql = Sql & ChronoAnnee & ".Année, " & ChronoAnnee & ".[Clé Ch], "
Sql = Sql & " " & ChronoAnnee & ".rv, " & ChronoAnnee & ".Rév "
Sql = Sql & "FROM " & ChronoAnnee & " "

Sql = "SELECT Agent_2.[Nom ag] AS Red_Nom,  "
Sql = Sql & "Agent_2.[Prénom ag] AS Red_P_Nom,  "
Sql = Sql & "Agent_1.[Nom ag] AS Verif_Nom,  "
Sql = Sql & "Agent_1.[Prénom ag] AS Verif_P_Nom,  "
Sql = Sql & "Agent.[Nom ag] AS Apr_Nom,  "
Sql = Sql & "Agent.[Prénom ag] AS Apr_P_Nom,  "
Sql = Sql & "" & ChronoAnnee & ".[Clé ty], " & ChronoAnnee & ".[Clé ac],  "
Sql = Sql & "" & ChronoAnnee & ".Année, " & ChronoAnnee & ".[Clé Ch],  "
Sql = Sql & "" & ChronoAnnee & ".rv, " & ChronoAnnee & ".Rév  "
Sql = Sql & "FROM ((" & ChronoAnnee & " INNER JOIN Agent ON " & ChronoAnnee & ".[Clé ap] = Agent.[Clé ag])  "
Sql = Sql & "INNER JOIN Agent AS Agent_1 ON " & ChronoAnnee & ".[Clé ve] = Agent_1.[Clé ag])  "
Sql = Sql & "INNER JOIN Agent AS Agent_2 ON " & ChronoAnnee & ".[Clé re] = Agent_2.[Clé ag] "
Sql = Sql & "WHERE " & ChronoAnnee & ".[Clé ac] = " & txt6.Text & " and " & ChronoAnnee & ".[Clé ty] = 'PL' "
Sql = Sql & "ORDER BY " & ChronoAnnee & ".[Clé Ch] DESC;"
Set RsBaseNum = ConNumPlan.OpenRecordSet(Sql)

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



Sql = "SELECT " & ChronoAnnee & ".[Clé ty], " & ChronoAnnee & ".[Clé ac],  "
Sql = Sql & ChronoAnnee & ".Année, " & ChronoAnnee & ".[Clé Ch], "
Sql = Sql & " " & ChronoAnnee & ".rv, " & ChronoAnnee & ".Rév "
Sql = Sql & "FROM " & ChronoAnnee & " "

Sql = "SELECT Agent_2.[Nom ag] AS Red_Nom,  "
Sql = Sql & "Agent_2.[Prénom ag] AS Red_P_Nom,  "
Sql = Sql & "Agent_1.[Nom ag] AS Verif_Nom,  "
Sql = Sql & "Agent_1.[Prénom ag] AS Verif_P_Nom,  "
Sql = Sql & "Agent.[Nom ag] AS Apr_Nom,  "
Sql = Sql & "Agent.[Prénom ag] AS Apr_P_Nom,  "
Sql = Sql & "" & ChronoAnnee & ".[Clé ty], " & ChronoAnnee & ".[Clé ac],  "
Sql = Sql & "" & ChronoAnnee & ".Année, " & ChronoAnnee & ".[Clé Ch],  "
Sql = Sql & "" & ChronoAnnee & ".rv, " & ChronoAnnee & ".Rév  "
Sql = Sql & "FROM ((" & ChronoAnnee & " INNER JOIN Agent ON " & ChronoAnnee & ".[Clé ap] = Agent.[Clé ag])  "
Sql = Sql & "INNER JOIN Agent AS Agent_1 ON " & ChronoAnnee & ".[Clé ve] = Agent_1.[Clé ag])  "
Sql = Sql & "INNER JOIN Agent AS Agent_2 ON " & ChronoAnnee & ".[Clé re] = Agent_2.[Clé ag] "
Sql = Sql & "WHERE " & ChronoAnnee & ".[Clé ac] = " & txt6.Text & " and " & ChronoAnnee & ".[Clé ty] = 'OU' "
Sql = Sql & "ORDER BY " & ChronoAnnee & ".[Clé Ch] DESC;"
Set RsBaseNum = ConNumPlan.OpenRecordSet(Sql)

While RsBaseNum.EOF = False
    txt9.AddItem "" & RsBaseNum![Clé ty] & "_" & RsBaseNum![Clé ac] & "_" & RsBaseNum![Année] & _
    "_" & RsBaseNum![Clé Ch] & "_" & RsBaseNum![rv] & "_" & RsBaseNum![Rév]
    txt9.List(txt9.ListCount - 1, 1) = "" & RsBaseNum![Clé Ch]
    txt9.List(txt9.ListCount - 1, 2) = "" & RsBaseNum![Red_Nom] & " " & RsBaseNum![Red_P_Nom]
    txt9.List(txt9.ListCount - 1, 3) = "" & RsBaseNum![Verif_Nom] & " " & RsBaseNum![Verif_P_Nom]
    txt9.List(txt9.ListCount - 1, 3) = "" & RsBaseNum![Apr_Nom] & " " & RsBaseNum![Apr_P_Nom]
    RsBaseNum.MoveNext
Wend
Sql = "SELECT " & ChronoAnnee & ".[Clé ty], " & ChronoAnnee & ".[Clé ac],  "
Sql = Sql & ChronoAnnee & ".Année, " & ChronoAnnee & ".[Clé Ch], "
Sql = Sql & " " & ChronoAnnee & ".rv, " & ChronoAnnee & ".Rév "
Sql = Sql & "FROM " & ChronoAnnee & " "

Sql = "SELECT Agent_2.[Nom ag] AS Red_Nom,  "
Sql = Sql & "Agent_2.[Prénom ag] AS Red_P_Nom,  "
Sql = Sql & "Agent_1.[Nom ag] AS Verif_Nom,  "
Sql = Sql & "Agent_1.[Prénom ag] AS Verif_P_Nom,  "
Sql = Sql & "Agent.[Nom ag] AS Apr_Nom,  "
Sql = Sql & "Agent.[Prénom ag] AS Apr_P_Nom,  "
Sql = Sql & "" & ChronoAnnee & ".[Clé ty], " & ChronoAnnee & ".[Clé ac],  "
Sql = Sql & "" & ChronoAnnee & ".Année, " & ChronoAnnee & ".[Clé Ch],  "
Sql = Sql & "" & ChronoAnnee & ".rv, " & ChronoAnnee & ".Rév  "
Sql = Sql & "FROM ((" & ChronoAnnee & " INNER JOIN Agent ON " & ChronoAnnee & ".[Clé ap] = Agent.[Clé ag])  "
Sql = Sql & "INNER JOIN Agent AS Agent_1 ON " & ChronoAnnee & ".[Clé ve] = Agent_1.[Clé ag])  "
Sql = Sql & "INNER JOIN Agent AS Agent_2 ON " & ChronoAnnee & ".[Clé re] = Agent_2.[Clé ag] "
Sql = Sql & "WHERE " & ChronoAnnee & ".[Clé ac] = " & txt6.Text & " and " & ChronoAnnee & ".[Clé ty] = 'LI' "
Sql = Sql & "ORDER BY " & ChronoAnnee & ".[Clé Ch] DESC;"
Set RsBaseNum = ConNumPlan.OpenRecordSet(Sql)

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
UserForm3.Show
Maj Me.txt1

End Sub

Private Sub CommandButton7_Click()
Dim Sql As String
Dim Rs As Recordset
NbTxt = 21
If ValideChampsTexte(Me, NbTxt) = False Then Exit Sub

  
Sql = "INSERT INTO T_indiceProjet ( [Indice],Id_Pieces,PI,PL,OU,  Li,  IdStatus,  "
Sql = Sql & "Client, Destinataire, Service, DessineDate, DessineNOM, VerifieDate,  "
Sql = Sql & "VerifieNom, ApprouveDate, ApprouveNom, Responsable, Vague, Equipement,  "
Sql = Sql & "Ensemble, CleAc,RefPF,RefP ) "
Sql = Sql & "values(  "
'sql = sql & "'" & MyReplace(txt17) & "',  "
Sql = Sql & "'" & MyReplace(txt7.List(txt7.ListIndex, 5)) & "' "
Sql = Sql & "," & IdPieces & ", "
Sql = Sql & "'" & MyReplace(txt7) & "',  "
Sql = Sql & "'" & MyReplace(txt8) & "',  "
Sql = Sql & "'" & MyReplace(txt9) & "',  "
Sql = Sql & "'" & MyReplace(txt10) & "', "
Sql = Sql & "1 ,  "
Sql = Sql & "'" & MyReplace(txt11.Text) & "',  "
Sql = Sql & "'" & MyReplace(txt12) & "',  "
Sql = Sql & "'" & MyReplace(txt13) & "',  "
Sql = Sql & MyReplaceDate(txt16) & ",  "
Sql = Sql & "'" & MyReplace(txt17) & "',  "
Sql = Sql & MyReplaceDate(txt18) & ",  "
Sql = Sql & "'" & MyReplace((txt19)) & "',  "
Sql = Sql & MyReplaceDate(txt20) & ",  "
Sql = Sql & "'" & MyReplace(txt21) & "',  "
Sql = Sql & "'" & MyReplace(txt4) & "',  "
Sql = Sql & "'" & MyReplace(txt2) & "',  "
Sql = Sql & "'" & MyReplace(txt3) & "',  "
'Sql = Sql & "'" & MyReplace(txt19) & "',  "
Sql = Sql & "'" & MyReplace(txt5) & "',  "
Sql = Sql & "'" & MyReplace(txt6) & "',  "
Sql = Sql & "'" & MyReplace(txt14) & "',  "
Sql = Sql & "'" & MyReplace(txt15) & "'); "
Con.Exequte Sql


ImportXls.Charge Me



End Sub

Private Sub UserForm_Activate()
Dim Rs As Recordset
Dim RsBaseNum As Recordset
Dim Sql As String
Dim indexClient As Long
Dim RqChronoAnnee As String
RqChronoAnne = "[Chrono Requête " & Format(Date, "yyyy]")
ChronoAnnee = "[Chrono" & Format(Date, "yyyy]")
Sql = "SELECT T_Clients.Client, T_Clients.Formulaire "
Sql = Sql & "FROM T_Clients "
Sql = Sql & "ORDER BY T_Clients.Client;"
Con.OpenConnetion db
ConNumPlan.OpenConnetion DbNumPlan
Set Rs = Con.OpenRecordSet(Sql)
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

Sql = "SELECT Activité.[Clé ac], Responsables.Nom, Responsables.Prénom, " & ChronoAnnee & ".Destinataire "
Sql = Sql & "FROM (Responsables RIGHT JOIN Activité ON Responsables.[Clé res] = Activité.[Clé re]) INNER JOIN " & ChronoAnnee & " ON Activité.[Clé ac] = " & ChronoAnnee & ".[Clé ac] "
Sql = Sql & "WHERE " & ChronoAnnee & ".[Clé ty]='PL' "
Sql = Sql & "OR " & ChronoAnnee & ".[Clé ty]='OU' "
Sql = Sql & "OR " & ChronoAnnee & ".[Clé ty]='LI' "
Sql = Sql & "OR " & ChronoAnnee & ".[Clé ty]='PI' "
Sql = Sql & "GROUP BY Activité.[Clé ac], Responsables.Nom, Responsables.Prénom, " & ChronoAnnee & ".Destinataire "
Sql = Sql & "ORDER BY Activité.[Clé ac] DESC;"

Sql = "SELECT Activité.[Clé ac], Activité.[Date ac], Activité.[Clé re],  "
Sql = Sql & "Activité.Client, Activité.Int, Activité.[Lib ac], Activité.[Obs ac],  "
Sql = Sql & "Activité.[Clé tyac], Activité.[St ac], Activité.[Clé pr],  "
Sql = Sql & "Activité.[Clé ca], Activité.[vid ac]  "
Sql = Sql & "FROM Activité INNER JOIN " & ChronoAnnee & " ON Activité.[Clé ac] = " & ChronoAnnee & ".[Clé ac] "
Sql = Sql & "WHERE " & ChronoAnnee & ".[Clé ty]='PL'  "
Sql = Sql & "OR " & ChronoAnnee & ".[Clé ty]='OU'  "
Sql = Sql & "OR " & ChronoAnnee & ".[Clé ty]='LI'  "
Sql = Sql & "OR " & ChronoAnnee & ".[Clé ty]='PI' "
Sql = Sql & "ORDER BY Activité.[Clé ac] DESC;"


Sql = "SELECT Activité.[Clé ac], Responsables.Nom, Responsables.Prénom, Activité.Int "
Sql = Sql & "FROM Responsables INNER JOIN (Activité INNER JOIN " & ChronoAnnee & "  "
Sql = Sql & "ON Activité.[Clé ac] = " & ChronoAnnee & ".[Clé ac])  "
Sql = Sql & "ON Responsables.[Clé res] = Activité.[Clé re] "
Sql = Sql & "WHERE " & ChronoAnnee & ".[Clé ty]='PL'  "
Sql = Sql & "Or " & ChronoAnnee & ".[Clé ty]='OU'  "
Sql = Sql & "Or " & ChronoAnnee & ".[Clé ty]='LI'  "
Sql = Sql & "Or " & ChronoAnnee & ".[Clé ty]='PI' "
Sql = Sql & "GROUP BY Activité.[Clé ac], Responsables.Nom, Responsables.Prénom, Activité.Int "
Sql = Sql & "ORDER BY Activité.[Clé ac] DESC;"



''
'SELECT " & ChronoAnnee & ".[Clé ac], " & ChronoAnnee & ".[Clé ty]
'FROM " & ChronoAnnee & "
'GROUP BY " & ChronoAnnee & ".[Clé ac], " & ChronoAnnee & ".[Clé ty]
''HAVING (((" & ChronoAnnee & ".[Clé ty]) = "PL")) Or (((" & ChronoAnnee & ".[Clé ty]) = "OU")) Or (((" & ChronoAnnee & ".[Clé ty]) = "LI")) Or (((" & ChronoAnnee & ".[Clé ty]) = "PI"))
'ORDER BY " & ChronoAnnee & ".[Clé ac] DESC;



Set RsBaseNum = ConNumPlan.OpenRecordSet(Sql)


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


