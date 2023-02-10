VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Creer 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Créer un nouveau plan de câblage:"
   ClientHeight    =   8040
   ClientLeft      =   30
   ClientTop       =   300
   ClientWidth     =   11670
   Icon            =   "Creer.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OleObjectBlob   =   "Creer.dsx":030A
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Creer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ChronoAnneeEnCours As String
Dim ChronoAnnee_M1 As String

Public IdProjet As Long
Public IdPieces As Long
Public IdIndiceProjet As Long
Public NbTxt As Long
Dim Noquite As Boolean
Dim Qui As String


Public Sub chargement(MyDroit As String)
Qui = MyDroit
Dim Rs As Recordset
Dim RsBaseNum As Recordset
Dim Sql As String
Dim indexClient As Long
Dim RqChronoAnneeEnCours As String
Dim RqChronoAnneeEnCours_M1 As String
Dim ErreurCon As Boolean
ConBaseNum.TYPEBASE = ADO_TYPEBASE
ConBaseNum.SERVER = ADO_SERVER
ConBaseNum.User = ADO_User
ConBaseNum.PassWord = ADO_PassWord
ConBaseNum.BASE = DbNumPlan


If ConBaseNum.OpenConnetion = True Then
Noquite = True
If Trim("" & BdDateTable) <> "" Then
    RqChronoAnne = "[Chrono Requête " & BdDateTable & "]"
    ChronoAnneeEnCours = "[Chrono" & BdDateTable & "]"
    ChronoAnnee_M1 = "[Chrono" & Val(BdDateTable) - 1 & "]"
Else
     RqChronoAnne = "[Chrono Requête " & Format(Date, "yyyy]")
     ChronoAnneeEnCours = "[Chrono" & Format(Date, "yyyy]")
     ChronoAnnee_M1 = "[Chrono" & Val(Format(Date, "yyyy")) - 1 & "]"
End If

Sql = "SELECT T_Clients.Client, T_Clients.Formulaire "
Sql = Sql & "FROM T_Clients "
Sql = Sql & "ORDER BY T_Clients.Client;"


Set Rs = Con.OpenRecordSet(Sql)
DoEvents
 Me.txt11.AddItem ""
While Rs.EOF = False
DoEvents
    Me.txt11.AddItem Trim("" & Rs!Client)
        If Me.txt11.ListCount = 1 Then Me.txt11.Text = Trim("" & Rs!Client)

    Rs.MoveNext
Wend






Sql = "SELECT Activité.[Clé ac], Responsables.Nom, Responsables.Prénom, Activité.Int "
Sql = Sql & "FROM (Responsables INNER JOIN (Activité LEFT JOIN " & ChronoAnneeEnCours & " ON Activité.[Clé ac] = "
Sql = Sql & " " & ChronoAnneeEnCours & ".[Clé ac]) ON Responsables.[Clé res] = Activité.[Clé re]) LEFT JOIN  "
Sql = Sql & ChronoAnnee_M1 & " ON Activité.[Clé ac] = " & ChronoAnnee_M1 & ".[Clé ac] "
Sql = Sql & "WHERE (" & ChronoAnneeEnCours & ".[Clé ty]='PL'  "
Sql = Sql & "Or " & ChronoAnneeEnCours & ".[Clé ty]='OU'  "
Sql = Sql & "or " & ChronoAnneeEnCours & ".[Clé ty]='LI'  "
Sql = Sql & "Or " & ChronoAnneeEnCours & ".[Clé ty]='PI') "
Sql = Sql & "Or "
Sql = Sql & "(" & ChronoAnnee_M1 & ".[Clé ty]='PL'  "
Sql = Sql & "Or " & ChronoAnnee_M1 & ".[Clé ty]='OU'  "
Sql = Sql & "or " & ChronoAnnee_M1 & ".[Clé ty]='LI'  "
Sql = Sql & "Or " & ChronoAnnee_M1 & ".[Clé ty]='PI') "
Sql = Sql & "GROUP BY Activité.[Clé ac], Responsables.Nom, Responsables.Prénom, Activité.Int "
Sql = Sql & "ORDER BY Activité.[Clé ac] DESC;"



'Sql = "SELECT Activité.[Clé ac], Responsables.Nom, Responsables.Prénom, Activité.Int "
'Sql = Sql & "FROM Responsables INNER JOIN (Activité INNER JOIN " & ChronoAnneeEnCours & "  "
'Sql = Sql & "ON Activité.[Clé ac] = " & ChronoAnneeEnCours & ".[Clé ac])  "
'Sql = Sql & "ON Responsables.[Clé res] = Activité.[Clé re] "
'Sql = Sql & "WHERE " & ChronoAnneeEnCours & ".[Clé ty]='PL'  "
'Sql = Sql & "Or " & ChronoAnneeEnCours & ".[Clé ty]='OU'  "
'Sql = Sql & "or " & ChronoAnneeEnCours & ".[Clé ty]='LI'  "
'Sql = Sql & "Or " & ChronoAnneeEnCours & ".[Clé ty]='PI' "
'Sql = Sql & "GROUP BY Activité.[Clé ac], Responsables.Nom, Responsables.Prénom, Activité.Int "
'Sql = Sql & "ORDER BY Activité.[Clé ac] DESC;"


 
' Sql = "SELECT Activité.[Clé ac], Responsables.Nom, Responsables.Prénom, Activité.Int "
'Sql = Sql & "FROM (Responsables INNER JOIN (Activité LEFT JOIN  " & ChronoAnneeEnCours & "  ON Activité.[Clé ac] =  "
'Sql = Sql & " " & ChronoAnneeEnCours & " .[Clé ac]) ON Responsables.[Clé res] = Activité.[Clé re]) LEFT  "
'Sql = Sql & "JOIN  " & ChronoAnnee_M1 & "  ON Activité.[Clé ac] =  " & ChronoAnnee_M1 & " .[Clé ac] "
'Sql = Sql & "WHERE ( " & ChronoAnneeEnCours & " .[Clé ty]='PL' Or  " & ChronoAnneeEnCours & " .[Clé ty]='OU' Or  "
'Sql = Sql & " " & ChronoAnneeEnCours & " .[Clé ty]='LI' Or  " & ChronoAnneeEnCours & " .[Clé ty])='PI'))  "
'Sql = Sql & "OR ((( " & ChronoAnnee_M1 & " .[Clé ty])='PL' Or ( " & ChronoAnnee_M1 & " .[Clé ty])='OU'  "
'Sql = Sql & "Or ( " & ChronoAnnee_M1 & " .[Clé ty])='LI' Or ( " & ChronoAnnee_M1 & " .[Clé ty])='PI')) "
'Sql = Sql & "GROUP BY Activité.[Clé ac], Responsables.Nom, Responsables.Prénom, Activité.Int "
'Sql = Sql & "ORDER BY Activité.[Clé ac] DESC;"
'
'
 
 

Set RsBaseNum = ConBaseNum.OpenRecordSet(Sql)

DoEvents
While RsBaseNum.EOF = False
DoEvents
    txt6.AddItem CStr(RsBaseNum![Clé ac])
    txt6.List(txt6.ListCount - 1, 1) = Trim("" & RsBaseNum!Nom) & " " & Trim("" & RsBaseNum!Prénom)
      txt6.List(txt6.ListCount - 1, 2) = Trim("" & RsBaseNum!Int)
 
    RsBaseNum.MoveNext
Wend

Sql = "SELECT T_Liste_Projet.Projet FROM T_Liste_Projet ORDER BY T_Liste_Projet.Projet;"
Set Rs = Con.OpenRecordSet(Sql)
txt1.Clear
txt24.Clear
txt24.AddItem ""
While Rs.EOF = False
    txt1.AddItem Trim("" & Rs!Projet)
    txt24.AddItem Trim("" & Rs!Projet)
    Rs.MoveNext
Wend
Set Rs = Con.CloseRecordSet(Rs)
ConBaseNum.CloseConnection
DoEvents
Me.Show vbModal
Else
 MsgBox "Impossible de se connecter à la base de données : " & vbCrLf & DbNumPlan & vbCrLf & vbCrLf & "Vérifiez qu'elle n'est pas en cours d'utilisation ?" & vbCrLf & "Ou contactez votre Administrateur Réseaux.", vbCritical
End If
End Sub
Private Sub CommandButton10_Click()
Dim Sql As String
Dim Rs As Recordset
If Trim("" & txt6) = "" Then
    MsgBox "Vous devez sélectionner un N° d'Affaire.", vbExclamation
    txt6.SetFocus
    Exit Sub
End If
CherchPices.Charge Me, " LiAutoCadSave <>  Null AND Pere=0 and CleAc=" & Val(txt6) & "  and IdStatus<>4 and Archiver=false", True
Unload CherchPices

Sql = "SELECT T_indiceProjet.Id, [PI] & '_' & [PI_Indice] AS Piece FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Val(Me.Tag) & ";"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = False Then
Me.Pere = Rs!Piece
End If
If Trim("" & Me.Tag) <> "" Then
Sql = "SELECT SelectProjets.Projet, SelectProjets.Vague, SelectProjets.Equipement,  "
Sql = Sql & "SelectProjets.Ensemble, SelectProjets.Client,T_Pieces.BaseVehicule  "
Sql = Sql & "FROM SelectProjets LEFT JOIN (T_Pieces RIGHT JOIN T_indiceProjet   "
Sql = Sql & "ON T_Pieces.Id = T_indiceProjet.Id_Pieces) ON SelectProjets.Id = T_indiceProjet.Id  "
Sql = Sql & "WHERE SelectProjets.Id=" & Me.Tag & ";"
Set Rs = Con.OpenRecordSet(Sql)
For I = 0 To txt1.ListCount - 1
    If txt1.List(I) = "" & Rs!Projet Then Me.txt1.ListIndex = I: Exit For
Next

  
     txt2 = "" & Rs!Vague

    txt3 = "" & Rs!Equipement


     txt5 = "" & Rs!Ensemble

For I = 0 To txt24.ListCount - 1
    If txt24.List(I) = "" & Rs!BaseVehicule Then Me.txt24.ListIndex = I: Exit For
Next
For I = 0 To txt11.ListCount - 1
    If txt11.List(I) = "" & Rs!Client Then Me.txt11.ListIndex = I: Exit For
Next
End If
Set Rs = Con.CloseRecordSet(Rs)

Me.Pere.Tag = Val(Me.Tag)
Set Rs = Con.CloseRecordSet(Rs)
MajFils
End Sub
Sub MajFils()
Dim Rs As Recordset
Dim Sql As String


Sql = "SELECT T_indiceProjet.PL, T_indiceProjet.PL_Indice,  "
Sql = Sql & "T_indiceProjet.[OU], T_indiceProjet.OU_Indice,  "
Sql = Sql & "T_indiceProjet.Li, T_indiceProjet.LI_Indice, T_indiceProjet.CleAc "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Val(Me.Tag) & ";"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = False Then
MajListe txt8, Rs, "PL"
MajListe txt9, Rs, "OU"
MajListe txt10, Rs, "LI"
End If
Set Rs = Con.CloseRecordSet(Rs)
End Sub
Sub MajListe(MyListe As Object, Rs As Recordset, Mytype As String)
For I = 0 To MyListe.ListCount - 1
Debug.Print UCase(Trim(MyListe.List(I, 0))) & " : " & UCase(Trim("" & Rs(Mytype)) & "_Rév.:_" & Trim("" & Rs(Mytype & "_indice")))
If UCase(Trim(MyListe.List(I, 0))) = UCase(Trim("" & Rs(Mytype)) & "_Rév.:_" & Trim("" & Rs(Mytype & "_indice"))) Then
    MyListe.ListIndex = I
    Exit For
End If
Next

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






Private Sub txt17_Change()
If Trim("" & Me.txt17.Text) = "" Then
     Me.txt16 = ""
Else
    Me.txt16 = Format(Date, "dd/mm/yyyy")
End If


End Sub



Private Sub txt6_Click()
Dim Sql As String
Dim RsBaseNum As Recordset
ConBaseNum.TYPEBASE = ADO_TYPEBASE
ConBaseNum.SERVER = ADO_SERVER
ConBaseNum.User = ADO_User
ConBaseNum.PassWord = ADO_PassWord
ConBaseNum.BASE = DbNumPlan


If ConBaseNum.OpenConnetion = True Then

txt7.Clear
txt8.Clear
txt9.Clear
txt10.Clear


Sql = "SELECT Agent_2.[Nom ag] AS Red_Nom,  "
Sql = Sql & "Agent_2.[Prénom ag] AS Red_P_Nom,  "
Sql = Sql & "Agent_1.[Nom ag] AS Verif_Nom,  "
Sql = Sql & "Agent_1.[Prénom ag] AS Verif_P_Nom,  "
Sql = Sql & "Agent.[Nom ag] AS Apr_Nom,  "
Sql = Sql & "Agent.[Prénom ag] AS Apr_P_Nom,  "
Sql = Sql & "" & ChronoAnneeEnCours & ".[Clé ty], " & ChronoAnneeEnCours & ".[Clé ac],  "
Sql = Sql & "" & ChronoAnneeEnCours & ".Année, " & ChronoAnneeEnCours & ".[Clé Ch],  "
Sql = Sql & "" & ChronoAnneeEnCours & ".rv, " & ChronoAnneeEnCours & ".Rév  "
Sql = Sql & "FROM ((" & ChronoAnneeEnCours & " INNER JOIN Agent ON " & ChronoAnneeEnCours & ".[Clé ap] = Agent.[Clé ag])  "
Sql = Sql & "INNER JOIN Agent AS Agent_1 ON " & ChronoAnneeEnCours & ".[Clé ve] = Agent_1.[Clé ag])  "
Sql = Sql & "INNER JOIN Agent AS Agent_2 ON " & ChronoAnneeEnCours & ".[Clé re] = Agent_2.[Clé ag] "
Sql = Sql & "WHERE " & ChronoAnneeEnCours & ".[Clé ac] = " & txt6.Text & " and " & ChronoAnneeEnCours & ".[Clé ty] = 'PI' "
Sql = Sql & "ORDER BY " & ChronoAnneeEnCours & ".[Clé Ch] DESC;"

Set RsBaseNum = ConBaseNum.OpenRecordSet(Sql)

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


Sql = "SELECT Agent_2.[Nom ag] AS Red_Nom,  "
Sql = Sql & "Agent_2.[Prénom ag] AS Red_P_Nom,  "
Sql = Sql & "Agent_1.[Nom ag] AS Verif_Nom,  "
Sql = Sql & "Agent_1.[Prénom ag] AS Verif_P_Nom,  "
Sql = Sql & "Agent.[Nom ag] AS Apr_Nom,  "
Sql = Sql & "Agent.[Prénom ag] AS Apr_P_Nom,  "
Sql = Sql & "" & ChronoAnnee_M1 & ".[Clé ty], " & ChronoAnnee_M1 & ".[Clé ac],  "
Sql = Sql & "" & ChronoAnnee_M1 & ".Année, " & ChronoAnnee_M1 & ".[Clé Ch],  "
Sql = Sql & "" & ChronoAnnee_M1 & ".rv, " & ChronoAnnee_M1 & ".Rév  "
Sql = Sql & "FROM ((" & ChronoAnnee_M1 & " INNER JOIN Agent ON " & ChronoAnnee_M1 & ".[Clé ap] = Agent.[Clé ag])  "
Sql = Sql & "INNER JOIN Agent AS Agent_1 ON " & ChronoAnnee_M1 & ".[Clé ve] = Agent_1.[Clé ag])  "
Sql = Sql & "INNER JOIN Agent AS Agent_2 ON " & ChronoAnnee_M1 & ".[Clé re] = Agent_2.[Clé ag] "
Sql = Sql & "WHERE " & ChronoAnnee_M1 & ".[Clé ac] = " & txt6.Text & " and " & ChronoAnnee_M1 & ".[Clé ty] = 'PI' "
Sql = Sql & "ORDER BY " & ChronoAnnee_M1 & ".[Clé Ch] DESC;"
Set RsBaseNum = ConBaseNum.OpenRecordSet(Sql)

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



Sql = "SELECT Agent_2.[Nom ag] AS Red_Nom,  "
Sql = Sql & "Agent_2.[Prénom ag] AS Red_P_Nom,  "
Sql = Sql & "Agent_1.[Nom ag] AS Verif_Nom,  "
Sql = Sql & "Agent_1.[Prénom ag] AS Verif_P_Nom,  "
Sql = Sql & "Agent.[Nom ag] AS Apr_Nom,  "
Sql = Sql & "Agent.[Prénom ag] AS Apr_P_Nom,  "
Sql = Sql & "" & ChronoAnneeEnCours & ".[Clé ty], " & ChronoAnneeEnCours & ".[Clé ac],  "
Sql = Sql & "" & ChronoAnneeEnCours & ".Année, " & ChronoAnneeEnCours & ".[Clé Ch],  "
Sql = Sql & "" & ChronoAnneeEnCours & ".rv, " & ChronoAnneeEnCours & ".Rév  "
Sql = Sql & "FROM ((" & ChronoAnneeEnCours & " INNER JOIN Agent ON " & ChronoAnneeEnCours & ".[Clé ap] = Agent.[Clé ag])  "
Sql = Sql & "INNER JOIN Agent AS Agent_1 ON " & ChronoAnneeEnCours & ".[Clé ve] = Agent_1.[Clé ag])  "
Sql = Sql & "INNER JOIN Agent AS Agent_2 ON " & ChronoAnneeEnCours & ".[Clé re] = Agent_2.[Clé ag] "
Sql = Sql & "WHERE " & ChronoAnneeEnCours & ".[Clé ac] = " & txt6.Text & " and " & ChronoAnneeEnCours & ".[Clé ty] = 'PL' "
Sql = Sql & "ORDER BY " & ChronoAnneeEnCours & ".[Clé Ch] DESC;"
Set RsBaseNum = ConBaseNum.OpenRecordSet(Sql)

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



Sql = "SELECT Agent_2.[Nom ag] AS Red_Nom,  "
Sql = Sql & "Agent_2.[Prénom ag] AS Red_P_Nom,  "
Sql = Sql & "Agent_1.[Nom ag] AS Verif_Nom,  "
Sql = Sql & "Agent_1.[Prénom ag] AS Verif_P_Nom,  "
Sql = Sql & "Agent.[Nom ag] AS Apr_Nom,  "
Sql = Sql & "Agent.[Prénom ag] AS Apr_P_Nom,  "
Sql = Sql & "" & ChronoAnnee_M1 & ".[Clé ty], " & ChronoAnnee_M1 & ".[Clé ac],  "
Sql = Sql & "" & ChronoAnnee_M1 & ".Année, " & ChronoAnnee_M1 & ".[Clé Ch],  "
Sql = Sql & "" & ChronoAnnee_M1 & ".rv, " & ChronoAnnee_M1 & ".Rév  "
Sql = Sql & "FROM ((" & ChronoAnnee_M1 & " INNER JOIN Agent ON " & ChronoAnnee_M1 & ".[Clé ap] = Agent.[Clé ag])  "
Sql = Sql & "INNER JOIN Agent AS Agent_1 ON " & ChronoAnnee_M1 & ".[Clé ve] = Agent_1.[Clé ag])  "
Sql = Sql & "INNER JOIN Agent AS Agent_2 ON " & ChronoAnnee_M1 & ".[Clé re] = Agent_2.[Clé ag] "
Sql = Sql & "WHERE " & ChronoAnnee_M1 & ".[Clé ac] = " & txt6.Text & " and " & ChronoAnnee_M1 & ".[Clé ty] = 'PL' "
Sql = Sql & "ORDER BY " & ChronoAnnee_M1 & ".[Clé Ch] DESC;"
Set RsBaseNum = ConBaseNum.OpenRecordSet(Sql)

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





Sql = "SELECT Agent_2.[Nom ag] AS Red_Nom,  "
Sql = Sql & "Agent_2.[Prénom ag] AS Red_P_Nom,  "
Sql = Sql & "Agent_1.[Nom ag] AS Verif_Nom,  "
Sql = Sql & "Agent_1.[Prénom ag] AS Verif_P_Nom,  "
Sql = Sql & "Agent.[Nom ag] AS Apr_Nom,  "
Sql = Sql & "Agent.[Prénom ag] AS Apr_P_Nom,  "
Sql = Sql & "" & ChronoAnneeEnCours & ".[Clé ty], " & ChronoAnneeEnCours & ".[Clé ac],  "
Sql = Sql & "" & ChronoAnneeEnCours & ".Année, " & ChronoAnneeEnCours & ".[Clé Ch],  "
Sql = Sql & "" & ChronoAnneeEnCours & ".rv, " & ChronoAnneeEnCours & ".Rév  "
Sql = Sql & "FROM ((" & ChronoAnneeEnCours & " INNER JOIN Agent ON " & ChronoAnneeEnCours & ".[Clé ap] = Agent.[Clé ag])  "
Sql = Sql & "INNER JOIN Agent AS Agent_1 ON " & ChronoAnneeEnCours & ".[Clé ve] = Agent_1.[Clé ag])  "
Sql = Sql & "INNER JOIN Agent AS Agent_2 ON " & ChronoAnneeEnCours & ".[Clé re] = Agent_2.[Clé ag] "
Sql = Sql & "WHERE " & ChronoAnneeEnCours & ".[Clé ac] = " & txt6.Text & " and " & ChronoAnneeEnCours & ".[Clé ty] = 'OU' "
Sql = Sql & "ORDER BY " & ChronoAnneeEnCours & ".[Clé Ch] DESC;"
Set RsBaseNum = ConBaseNum.OpenRecordSet(Sql)

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


Sql = "SELECT Agent_2.[Nom ag] AS Red_Nom,  "
Sql = Sql & "Agent_2.[Prénom ag] AS Red_P_Nom,  "
Sql = Sql & "Agent_1.[Nom ag] AS Verif_Nom,  "
Sql = Sql & "Agent_1.[Prénom ag] AS Verif_P_Nom,  "
Sql = Sql & "Agent.[Nom ag] AS Apr_Nom,  "
Sql = Sql & "Agent.[Prénom ag] AS Apr_P_Nom,  "
Sql = Sql & "" & ChronoAnnee_M1 & ".[Clé ty], " & ChronoAnnee_M1 & ".[Clé ac],  "
Sql = Sql & "" & ChronoAnnee_M1 & ".Année, " & ChronoAnnee_M1 & ".[Clé Ch],  "
Sql = Sql & "" & ChronoAnnee_M1 & ".rv, " & ChronoAnnee_M1 & ".Rév  "
Sql = Sql & "FROM ((" & ChronoAnnee_M1 & " INNER JOIN Agent ON " & ChronoAnnee_M1 & ".[Clé ap] = Agent.[Clé ag])  "
Sql = Sql & "INNER JOIN Agent AS Agent_1 ON " & ChronoAnnee_M1 & ".[Clé ve] = Agent_1.[Clé ag])  "
Sql = Sql & "INNER JOIN Agent AS Agent_2 ON " & ChronoAnnee_M1 & ".[Clé re] = Agent_2.[Clé ag] "
Sql = Sql & "WHERE " & ChronoAnnee_M1 & ".[Clé ac] = " & txt6.Text & " and " & ChronoAnnee_M1 & ".[Clé ty] = 'OU' "
Sql = Sql & "ORDER BY " & ChronoAnnee_M1 & ".[Clé Ch] DESC;"
Set RsBaseNum = ConBaseNum.OpenRecordSet(Sql)

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





Sql = "SELECT Agent_2.[Nom ag] AS Red_Nom,  "
Sql = Sql & "Agent_2.[Prénom ag] AS Red_P_Nom,  "
Sql = Sql & "Agent_1.[Nom ag] AS Verif_Nom,  "
Sql = Sql & "Agent_1.[Prénom ag] AS Verif_P_Nom,  "
Sql = Sql & "Agent.[Nom ag] AS Apr_Nom,  "
Sql = Sql & "Agent.[Prénom ag] AS Apr_P_Nom,  "
Sql = Sql & "" & ChronoAnneeEnCours & ".[Clé ty], " & ChronoAnneeEnCours & ".[Clé ac],  "
Sql = Sql & "" & ChronoAnneeEnCours & ".Année, " & ChronoAnneeEnCours & ".[Clé Ch],  "
Sql = Sql & "" & ChronoAnneeEnCours & ".rv, " & ChronoAnneeEnCours & ".Rév  "
Sql = Sql & "FROM ((" & ChronoAnneeEnCours & " INNER JOIN Agent ON " & ChronoAnneeEnCours & ".[Clé ap] = Agent.[Clé ag])  "
Sql = Sql & "INNER JOIN Agent AS Agent_1 ON " & ChronoAnneeEnCours & ".[Clé ve] = Agent_1.[Clé ag])  "
Sql = Sql & "INNER JOIN Agent AS Agent_2 ON " & ChronoAnneeEnCours & ".[Clé re] = Agent_2.[Clé ag] "
Sql = Sql & "WHERE " & ChronoAnneeEnCours & ".[Clé ac] = " & txt6.Text & " and " & ChronoAnneeEnCours & ".[Clé ty] = 'LI' "
Sql = Sql & "ORDER BY " & ChronoAnneeEnCours & ".[Clé Ch] DESC;"
Set RsBaseNum = ConBaseNum.OpenRecordSet(Sql)

While RsBaseNum.EOF = False
    txt10.AddItem "" & RsBaseNum![Clé ty] & "_" & RsBaseNum![Clé ac] & "_" & RsBaseNum![Année] & _
    "_" & RsBaseNum![Clé Ch] & "_" & RsBaseNum![rv] & "_" & RsBaseNum![Rév]
    txt10.List(txt10.ListCount - 1, 1) = "" & RsBaseNum![Clé Ch]
     txt10.List(txt10.ListCount - 1, 2) = " " & RsBaseNum![Rév]
    RsBaseNum.MoveNext
Wend



Sql = "SELECT Agent_2.[Nom ag] AS Red_Nom,  "
Sql = Sql & "Agent_2.[Prénom ag] AS Red_P_Nom,  "
Sql = Sql & "Agent_1.[Nom ag] AS Verif_Nom,  "
Sql = Sql & "Agent_1.[Prénom ag] AS Verif_P_Nom,  "
Sql = Sql & "Agent.[Nom ag] AS Apr_Nom,  "
Sql = Sql & "Agent.[Prénom ag] AS Apr_P_Nom,  "
Sql = Sql & "" & ChronoAnnee_M1 & ".[Clé ty], " & ChronoAnnee_M1 & ".[Clé ac],  "
Sql = Sql & "" & ChronoAnnee_M1 & ".Année, " & ChronoAnnee_M1 & ".[Clé Ch],  "
Sql = Sql & "" & ChronoAnnee_M1 & ".rv, " & ChronoAnnee_M1 & ".Rév  "
Sql = Sql & "FROM ((" & ChronoAnnee_M1 & " INNER JOIN Agent ON " & ChronoAnnee_M1 & ".[Clé ap] = Agent.[Clé ag])  "
Sql = Sql & "INNER JOIN Agent AS Agent_1 ON " & ChronoAnnee_M1 & ".[Clé ve] = Agent_1.[Clé ag])  "
Sql = Sql & "INNER JOIN Agent AS Agent_2 ON " & ChronoAnnee_M1 & ".[Clé re] = Agent_2.[Clé ag] "
Sql = Sql & "WHERE " & ChronoAnnee_M1 & ".[Clé ac] = " & txt6.Text & " and " & ChronoAnnee_M1 & ".[Clé ty] = 'LI' "
Sql = Sql & "ORDER BY " & ChronoAnnee_M1 & ".[Clé Ch] DESC;"
Set RsBaseNum = ConBaseNum.OpenRecordSet(Sql)

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
Else
 MsgBox "Impossible de se connecter à la base de données : " & vbCrLf & DbNumPlan & vbCrLf & vbCrLf & "Vérifiez qu'elle n'est pas en cours d'utilisation ?" & vbCrLf & "Ou contactez votre Administrateur Réseaux.", vbCritical
 Me.Hide
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
UserForm1.charger txt5, vbCrLf, "Ensemble:"

End Sub

Private Sub CommandButton4_Click()
UserForm1.charger txt3, ";", "Equipement:", "_"
End Sub

Private Sub CommandButton5_Click()
UserForm1.charger txt2, " ", "Vagues:", " "
End Sub

Private Sub CommandButton6_Click()
UserForm3.Show vbModal
Unload UserForm3
Maj Me.txt11

End Sub

Private Sub CommandButton7_Click()
Dim Sql As String
Dim Rs As Recordset
Dim Pose As Long
Dim Txt As String

NbTxt = 25
If ValideChampsTexte(Me, NbTxt) = False Then Exit Sub

Sql = "SELECT T_Projet.id FROM T_Projet "


Sql = Sql & "WHERE T_Projet.Projet='" & MyReplace(txt1) & "';"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = True Then

    Sql = "INSERT INTO T_Projet ( Projet ) "
    Sql = Sql & "VALUES( '" & MyReplace(txt1) & "');"
    Con.Execute Sql
End If
Rs.Requery

IdProjet = Rs!Id

Pose = InStr(1, txt7, "_Rév.:_")
If Pose = 0 Then
    Txt = txt7
Else
    Txt = Mid(txt7, 1, Pose - 1)
End If
'txt = Replace(txt7, "_Rév.:", "")
Sql = "SELECT T_Pieces.Id, T_Projet.Projet "
Sql = Sql & "FROM T_Projet INNER JOIN T_Pieces ON T_Projet.id = T_Pieces.IdProjet "
Sql = Sql & "WHERE T_Pieces.Description='" & MyReplace(Txt) & "';"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = False Then
    MsgBox "La pièce : " & Txt & " existe déjà dans le projet : " & Rs!Projet & vbCrLf & "Opération d'ajout annulée"
    Exit Sub
End If
Sql = "INSERT INTO T_Pieces ( IdProjet, Description,BaseVehicule)"
Sql = Sql & "VALUES( " & IdProjet & ", '" & MyReplace(Txt) & "', '" & MyReplace(txt24) & "');"
 Con.Execute Sql
 Rs.Requery
 IdPieces = Rs!Id
 Sql = "INSERT INTO T_indiceProjet ( Id_Pieces,Ref_Piece_CLI,Ref_Plan_CLI,Ref_PF, Vague, Equipement, Responsable, Ensemble, "
 Sql = Sql & "CleAc, PI, PL,PL_Indice, OU,OU_Indice, Li,LI_Indice, Client, Destinataire, Service, RefPF, RefP, DessineDate,  "
 Sql = Sql & "DessineNOM, VerifieDate, VerifieNom, ApprouveDate, ApprouveNom ,PI_Indice,Masse "
 
 If Trim("" & Me.Pere) <> "" Then
    Sql = Sql & ",pere "
 End If
 
 Sql = Sql & ",NbCartouche,RefPieceClient) "
 
  Pose = InStr(1, txt7, "_Rév.:_")
If Pose = 0 Then
    Txt = txt7
Else
    Txt = Mid(txt7, 1, Pose - 1)
End If
 Sql = Sql & "VALUES ( " & IdPieces & " ,'" & MyReplace(Ref_Piece_CLI) & "','" & MyReplace(Ref_Plan_CLI) & "','" & MyReplace(Ref_PF) & "', '"
 Sql = Sql & MyReplace(txt2) & "' , '" & MyReplace(txt3) & "', "
 Sql = Sql & "'" & MyReplace(txt4) & "' , '" & MyReplace(txt5) & "' , '" & MyReplace(txt6) & "' ,  "
 Sql = Sql & "'" & MyReplace(Txt) & "' , "
 Pose = InStr(1, txt8, "_Rév.:_")
If Pose = 0 Then
    Txt = txt8
Else
    Txt = Mid(txt8, 1, Pose - 1)
End If
Sql = Sql & "'" & MyReplace(Txt) & "' ,"
Sql = Sql & "'" & MyReplace(txt8.List(txt8.ListIndex, 5)) & "',"

 Pose = InStr(1, txt9, "_Rév.:_")
If Pose = 0 Then
    Txt = txt9
Else
    Txt = Mid(txt9, 1, Pose - 1)
End If
 Sql = Sql & "'" & MyReplace(Txt) & "' ,  "
 Sql = Sql & "'" & MyReplace(txt9.List(txt9.ListIndex, 5)) & "',"

 Pose = InStr(1, txt10, "_Rév.:_")
If Pose = 0 Then
    Txt = txt10
Else
    Txt = Mid(txt10, 1, Pose - 1)
End If
Sql = Sql & "'" & MyReplace(Txt) & "', "
Sql = Sql & "'" & MyReplace(txt10.List(txt10.ListIndex, 2)) & "',"


 Sql = Sql & "'" & MyReplace(txt11) & "' , '" & MyReplace(txt12) & "' ,  "
 Sql = Sql & "'" & MyReplace(txt13) & "' , '" & MyReplace(txt14) & "' , '" & MyReplace(txt15) & "' ,  "
 Sql = Sql & "" & MyReplaceDate(txt16) & " , '" & MyReplace(txt17) & "' , " & MyReplaceDate(txt18) & " ,  "
 Sql = Sql & "'" & MyReplace(txt19) & "' , " & MyReplaceDate(txt20) & " , '" & MyReplace(txt21) & "','" & txt7.List(txt7.ListIndex, 5) & "','" & TXT25 & "' "
 
 If Trim("" & Me.Pere) <> "" Then
    Sql = Sql & "," & Me.Pere.Tag
 End If
 
 Sql = Sql & "," & txt22 & ","
  Sql = Sql & "'" & MyReplace(txt23) & "');"
Con.Execute Sql
Sql = "SELECT T_indiceProjet.Id "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id_Pieces=" & IdPieces & " ;"
Set Rs = Con.OpenRecordSet(Sql)
IdIndiceProjet = Rs!Id
NbTxt = 21
Noquite = False
Me.Hide
If Trim("" & Me.Pere) <> "" Then
DoEvents

Modifier.Charge Me, Qui
Unload Modifier
Else
    Noquite = False
    ImportXls.Charge Me, Qui
'    Unload ImportXls
End If
Noquite = False
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Cancel = Noquite
End Sub
