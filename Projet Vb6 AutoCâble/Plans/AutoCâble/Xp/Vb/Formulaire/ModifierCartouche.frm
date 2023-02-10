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

sql = "SELECT " & ChronoAnnee & ".[Cl� ty], " & ChronoAnnee & ".[Cl� ac],  "
sql = sql & ChronoAnnee & ".Ann�e, " & ChronoAnnee & ".[Cl� Ch], "
sql = sql & " " & ChronoAnnee & ".rv, " & ChronoAnnee & ".R�v "
sql = sql & "FROM " & ChronoAnnee & "  "

sql = "SELECT Agent_2.[Nom ag] AS Red_Nom,  "
sql = sql & "Agent_2.[Pr�nom ag] AS Red_P_Nom,  "
sql = sql & "Agent_1.[Nom ag] AS Verif_Nom,  "
sql = sql & "Agent_1.[Pr�nom ag] AS Verif_P_Nom,  "
sql = sql & "Agent.[Nom ag] AS Apr_Nom,  "
sql = sql & "Agent.[Pr�nom ag] AS Apr_P_Nom,  "
sql = sql & "" & ChronoAnnee & ".[Cl� ty], " & ChronoAnnee & ".[Cl� ac],  "
sql = sql & "" & ChronoAnnee & ".Ann�e, " & ChronoAnnee & ".[Cl� Ch],  "
sql = sql & "" & ChronoAnnee & ".rv, " & ChronoAnnee & ".R�v  "
sql = sql & "FROM ((" & ChronoAnnee & " INNER JOIN Agent ON " & ChronoAnnee & ".[Cl� ap] = Agent.[Cl� ag])  "
sql = sql & "INNER JOIN Agent AS Agent_1 ON " & ChronoAnnee & ".[Cl� ve] = Agent_1.[Cl� ag])  "
sql = sql & "INNER JOIN Agent AS Agent_2 ON " & ChronoAnnee & ".[Cl� re] = Agent_2.[Cl� ag] "

sql = sql & "WHERE " & ChronoAnnee & ".[Cl� ac] = " & txt6.Text & " and " & ChronoAnnee & ".[Cl� ty] = 'PI' "
sql = sql & "ORDER BY " & ChronoAnnee & ".[Cl� Ch] DESC;"
Set RsBaseNum = ConNumPlan.OpenRecordSet(sql)

While RsBaseNum.EOF = False

    txt7.AddItem "" & RsBaseNum![Cl� ty] & "_" & RsBaseNum![Cl� ac] & "_" & RsBaseNum![Ann�e] & _
    "_" & RsBaseNum![Cl� Ch] & "_" & RsBaseNum![rv] & "_" & RsBaseNum![R�v]
    txt7.List(txt7.ListCount - 1, 1) = "" & RsBaseNum![Cl� Ch]
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
    txt7.List(txt7.ListCount - 1, 5) = "" & RsBaseNum![R�v]

    RsBaseNum.MoveNext
Wend

sql = "SELECT " & ChronoAnnee & ".[Cl� ty], " & ChronoAnnee & ".[Cl� ac],  "
sql = sql & ChronoAnnee & ".Ann�e, " & ChronoAnnee & ".[Cl� Ch], "
sql = sql & " " & ChronoAnnee & ".rv, " & ChronoAnnee & ".R�v "
sql = sql & "FROM " & ChronoAnnee & " "

sql = "SELECT Agent_2.[Nom ag] AS Red_Nom,  "
sql = sql & "Agent_2.[Pr�nom ag] AS Red_P_Nom,  "
sql = sql & "Agent_1.[Nom ag] AS Verif_Nom,  "
sql = sql & "Agent_1.[Pr�nom ag] AS Verif_P_Nom,  "
sql = sql & "Agent.[Nom ag] AS Apr_Nom,  "
sql = sql & "Agent.[Pr�nom ag] AS Apr_P_Nom,  "
sql = sql & "" & ChronoAnnee & ".[Cl� ty], " & ChronoAnnee & ".[Cl� ac],  "
sql = sql & "" & ChronoAnnee & ".Ann�e, " & ChronoAnnee & ".[Cl� Ch],  "
sql = sql & "" & ChronoAnnee & ".rv, " & ChronoAnnee & ".R�v  "
sql = sql & "FROM ((" & ChronoAnnee & " INNER JOIN Agent ON " & ChronoAnnee & ".[Cl� ap] = Agent.[Cl� ag])  "
sql = sql & "INNER JOIN Agent AS Agent_1 ON " & ChronoAnnee & ".[Cl� ve] = Agent_1.[Cl� ag])  "
sql = sql & "INNER JOIN Agent AS Agent_2 ON " & ChronoAnnee & ".[Cl� re] = Agent_2.[Cl� ag] "
sql = sql & "WHERE " & ChronoAnnee & ".[Cl� ac] = " & txt6.Text & " and " & ChronoAnnee & ".[Cl� ty] = 'PL' "
sql = sql & "ORDER BY " & ChronoAnnee & ".[Cl� Ch] DESC;"
Set RsBaseNum = ConNumPlan.OpenRecordSet(sql)

While RsBaseNum.EOF = False
    txt8.AddItem "" & RsBaseNum![Cl� ty] & "_" & RsBaseNum![Cl� ac] & "_" & RsBaseNum![Ann�e] & _
    "_" & RsBaseNum![Cl� Ch] & "_" & RsBaseNum![rv] & "_" & RsBaseNum![R�v]
    txt8.List(txt8.ListCount - 1, 1) = "" & RsBaseNum![Cl� Ch]
    txt8.List(txt8.ListCount - 1, 1) = "" & RsBaseNum![Cl� Ch]
    txt8.List(txt8.ListCount - 1, 2) = "" & RsBaseNum![Red_Nom] & " " & RsBaseNum![Red_P_Nom]
    txt8.List(txt8.ListCount - 1, 3) = "" & RsBaseNum![Verif_Nom] & " " & RsBaseNum![Verif_P_Nom]
    txt8.List(txt8.ListCount - 1, 3) = "" & RsBaseNum![Apr_Nom] & " " & RsBaseNum![Apr_P_Nom]
    RsBaseNum.MoveNext
Wend



sql = "SELECT " & ChronoAnnee & ".[Cl� ty], " & ChronoAnnee & ".[Cl� ac],  "
sql = sql & ChronoAnnee & ".Ann�e, " & ChronoAnnee & ".[Cl� Ch], "
sql = sql & " " & ChronoAnnee & ".rv, " & ChronoAnnee & ".R�v "
sql = sql & "FROM " & ChronoAnnee & " "

sql = "SELECT Agent_2.[Nom ag] AS Red_Nom,  "
sql = sql & "Agent_2.[Pr�nom ag] AS Red_P_Nom,  "
sql = sql & "Agent_1.[Nom ag] AS Verif_Nom,  "
sql = sql & "Agent_1.[Pr�nom ag] AS Verif_P_Nom,  "
sql = sql & "Agent.[Nom ag] AS Apr_Nom,  "
sql = sql & "Agent.[Pr�nom ag] AS Apr_P_Nom,  "
sql = sql & "" & ChronoAnnee & ".[Cl� ty], " & ChronoAnnee & ".[Cl� ac],  "
sql = sql & "" & ChronoAnnee & ".Ann�e, " & ChronoAnnee & ".[Cl� Ch],  "
sql = sql & "" & ChronoAnnee & ".rv, " & ChronoAnnee & ".R�v  "
sql = sql & "FROM ((" & ChronoAnnee & " INNER JOIN Agent ON " & ChronoAnnee & ".[Cl� ap] = Agent.[Cl� ag])  "
sql = sql & "INNER JOIN Agent AS Agent_1 ON " & ChronoAnnee & ".[Cl� ve] = Agent_1.[Cl� ag])  "
sql = sql & "INNER JOIN Agent AS Agent_2 ON " & ChronoAnnee & ".[Cl� re] = Agent_2.[Cl� ag] "
sql = sql & "WHERE " & ChronoAnnee & ".[Cl� ac] = " & txt6.Text & " and " & ChronoAnnee & ".[Cl� ty] = 'OU' "
sql = sql & "ORDER BY " & ChronoAnnee & ".[Cl� Ch] DESC;"
Set RsBaseNum = ConNumPlan.OpenRecordSet(sql)

While RsBaseNum.EOF = False
    txt9.AddItem "" & RsBaseNum![Cl� ty] & "_" & RsBaseNum![Cl� ac] & "_" & RsBaseNum![Ann�e] & _
    "_" & RsBaseNum![Cl� Ch] & "_" & RsBaseNum![rv] & "_" & RsBaseNum![R�v]
    txt9.List(txt9.ListCount - 1, 1) = "" & RsBaseNum![Cl� Ch]
    txt9.List(txt9.ListCount - 1, 2) = "" & RsBaseNum![Red_Nom] & " " & RsBaseNum![Red_P_Nom]
    txt9.List(txt9.ListCount - 1, 3) = "" & RsBaseNum![Verif_Nom] & " " & RsBaseNum![Verif_P_Nom]
    txt9.List(txt9.ListCount - 1, 3) = "" & RsBaseNum![Apr_Nom] & " " & RsBaseNum![Apr_P_Nom]
    RsBaseNum.MoveNext
Wend
sql = "SELECT " & ChronoAnnee & ".[Cl� ty], " & ChronoAnnee & ".[Cl� ac],  "
sql = sql & ChronoAnnee & ".Ann�e, " & ChronoAnnee & ".[Cl� Ch], "
sql = sql & " " & ChronoAnnee & ".rv, " & ChronoAnnee & ".R�v "
sql = sql & "FROM " & ChronoAnnee & " "

sql = "SELECT Agent_2.[Nom ag] AS Red_Nom,  "
sql = sql & "Agent_2.[Pr�nom ag] AS Red_P_Nom,  "
sql = sql & "Agent_1.[Nom ag] AS Verif_Nom,  "
sql = sql & "Agent_1.[Pr�nom ag] AS Verif_P_Nom,  "
sql = sql & "Agent.[Nom ag] AS Apr_Nom,  "
sql = sql & "Agent.[Pr�nom ag] AS Apr_P_Nom,  "
sql = sql & "" & ChronoAnnee & ".[Cl� ty], " & ChronoAnnee & ".[Cl� ac],  "
sql = sql & "" & ChronoAnnee & ".Ann�e, " & ChronoAnnee & ".[Cl� Ch],  "
sql = sql & "" & ChronoAnnee & ".rv, " & ChronoAnnee & ".R�v  "
sql = sql & "FROM ((" & ChronoAnnee & " INNER JOIN Agent ON " & ChronoAnnee & ".[Cl� ap] = Agent.[Cl� ag])  "
sql = sql & "INNER JOIN Agent AS Agent_1 ON " & ChronoAnnee & ".[Cl� ve] = Agent_1.[Cl� ag])  "
sql = sql & "INNER JOIN Agent AS Agent_2 ON " & ChronoAnnee & ".[Cl� re] = Agent_2.[Cl� ag] "
sql = sql & "WHERE " & ChronoAnnee & ".[Cl� ac] = " & txt6.Text & " and " & ChronoAnnee & ".[Cl� ty] = 'LI' "
sql = sql & "ORDER BY " & ChronoAnnee & ".[Cl� Ch] DESC;"
Set RsBaseNum = ConNumPlan.OpenRecordSet(sql)

While RsBaseNum.EOF = False
    txt10.AddItem "" & RsBaseNum![Cl� ty] & "_" & RsBaseNum![Cl� ac] & "_" & RsBaseNum![Ann�e] & _
    "_" & RsBaseNum![Cl� Ch] & "_" & RsBaseNum![rv] & "_" & RsBaseNum![R�v]
    txt10.List(txt10.ListCount - 1, 1) = "" & RsBaseNum![Cl� Ch]
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
RqChronoAnne = "[Chrono Requ�te " & Format(Date, "yyyy]")
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




'Sql = "SELECT " & ChronoAnnee & ".[Cl� ty], " & ChronoAnnee & ".[Cl� ac],  "
'Sql = Sql & ChronoAnnee & ".Ann�e, " & ChronoAnnee & ".[Cl� Ch], "
'Sql = Sql & " " & ChronoAnnee & ".rv, " & ChronoAnnee & ".R�v "
'Sql = Sql & "FROM " & ChronoAnnee & " "
'Sql = Sql & "WHERE " & ChronoAnnee & ".[Cl� ty] = 'LI' "
'Sql = Sql & "ORDER BY " & ChronoAnnee & ".[Cl� Ch] DESC;"
'Set RsBaseNum = ConNumPlan.OpenRecordSet(Sql)
'
'While RsBaseNum.EOF = False
'    ComboBox1.AddItem "" & RsBaseNum![Cl� Ch]
'    RsBaseNum.MoveNext
'Wend

sql = "SELECT Activit�.[Cl� ac], Responsables.Nom, Responsables.Pr�nom, " & ChronoAnnee & ".Destinataire "
sql = sql & "FROM (Responsables RIGHT JOIN Activit� ON Responsables.[Cl� res] = Activit�.[Cl� re]) INNER JOIN " & ChronoAnnee & " ON Activit�.[Cl� ac] = " & ChronoAnnee & ".[Cl� ac] "
sql = sql & "WHERE " & ChronoAnnee & ".[Cl� ty]='PL' "
sql = sql & "OR " & ChronoAnnee & ".[Cl� ty]='OU' "
sql = sql & "OR " & ChronoAnnee & ".[Cl� ty]='LI' "
sql = sql & "OR " & ChronoAnnee & ".[Cl� ty]='PI' "
sql = sql & "GROUP BY Activit�.[Cl� ac], Responsables.Nom, Responsables.Pr�nom, " & ChronoAnnee & ".Destinataire "
sql = sql & "ORDER BY Activit�.[Cl� ac] DESC;"

sql = "SELECT Activit�.[Cl� ac], Activit�.[Date ac], Activit�.[Cl� re],  "
sql = sql & "Activit�.Client, Activit�.Int, Activit�.[Lib ac], Activit�.[Obs ac],  "
sql = sql & "Activit�.[Cl� tyac], Activit�.[St ac], Activit�.[Cl� pr],  "
sql = sql & "Activit�.[Cl� ca], Activit�.[vid ac]  "
sql = sql & "FROM Activit� INNER JOIN " & ChronoAnnee & " ON Activit�.[Cl� ac] = " & ChronoAnnee & ".[Cl� ac] "
sql = sql & "WHERE " & ChronoAnnee & ".[Cl� ty]='PL'  "
sql = sql & "OR " & ChronoAnnee & ".[Cl� ty]='OU'  "
sql = sql & "OR " & ChronoAnnee & ".[Cl� ty]='LI'  "
sql = sql & "OR " & ChronoAnnee & ".[Cl� ty]='PI' "
sql = sql & "ORDER BY Activit�.[Cl� ac] DESC;"


sql = "SELECT Activit�.[Cl� ac], Responsables.Nom, Responsables.Pr�nom, Activit�.Int "
sql = sql & "FROM Responsables INNER JOIN (Activit� INNER JOIN " & ChronoAnnee & "  "
sql = sql & "ON Activit�.[Cl� ac] = " & ChronoAnnee & ".[Cl� ac])  "
sql = sql & "ON Responsables.[Cl� res] = Activit�.[Cl� re] "
sql = sql & "WHERE " & ChronoAnnee & ".[Cl� ty]='PL'  "
sql = sql & "Or " & ChronoAnnee & ".[Cl� ty]='OU'  "
sql = sql & "Or " & ChronoAnnee & ".[Cl� ty]='LI'  "
sql = sql & "Or " & ChronoAnnee & ".[Cl� ty]='PI' "
sql = sql & "GROUP BY Activit�.[Cl� ac], Responsables.Nom, Responsables.Pr�nom, Activit�.Int "
sql = sql & "ORDER BY Activit�.[Cl� ac] DESC;"



''
'SELECT " & ChronoAnnee & ".[Cl� ac], " & ChronoAnnee & ".[Cl� ty]
'FROM " & ChronoAnnee & "
'GROUP BY " & ChronoAnnee & ".[Cl� ac], " & ChronoAnnee & ".[Cl� ty]
''HAVING (((" & ChronoAnnee & ".[Cl� ty]) = "PL")) Or (((" & ChronoAnnee & ".[Cl� ty]) = "OU")) Or (((" & ChronoAnnee & ".[Cl� ty]) = "LI")) Or (((" & ChronoAnnee & ".[Cl� ty]) = "PI"))
'ORDER BY " & ChronoAnnee & ".[Cl� ac] DESC;



Set RsBaseNum = ConNumPlan.OpenRecordSet(sql)


While RsBaseNum.EOF = False
    txt6.AddItem CStr(RsBaseNum![Cl� ac])
    txt6.List(txt6.ListCount - 1, 1) = Trim("" & RsBaseNum!NOM) & " " & Trim("" & RsBaseNum!Pr�nom)
      txt6.List(txt6.ListCount - 1, 2) = Trim("" & RsBaseNum!Int)
 
    RsBaseNum.MoveNext
Wend
Set Rs = Con.CloseRecordSet(Rs)
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Con.CloseConnection
ConNumPlan.CloseConnection
End Sub


