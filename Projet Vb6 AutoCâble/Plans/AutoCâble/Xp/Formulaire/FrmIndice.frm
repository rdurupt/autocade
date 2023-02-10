VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmIndice 
   Caption         =   "Changement d'indice :"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9150
   OleObjectBlob   =   "FrmIndice.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmIndice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CopieStrtxt7 As String
Dim CopieStrtxt8 As String
Dim CopieStrtxt9 As String
Dim CopieStrtxt10 As String
Dim boolExecute As Boolean
Dim Noquite As Boolean
Public Function Charge(Projet As String, Vague As String, Equipement As String, Ensemble As String, Client As String, Affaire As String, strtxt7 As String, strtxt8 As String, strtxt9 As String, strtxt10 As String) As Boolean
Dim sql As String
Dim RsBaseNum As Recordset
ConBaseNum.OpenConnetion DbNumPlan

txt1 = Projet
txt2 = Vague
txt3 = Equipement
txt4 = Ensemble
txt9 = Client
txt5.Clear

txt6.Clear
txt7.Clear
txt8.Clear
CopieStrtxt7 = strtxt7
 CopieStrtxt8 = strtxt8
 CopieStrtxt9 = strtxt9
 CopieStrtxt10 = strtxt10

RqChronoAnne = "[Chrono Requête " & Format(Date, "yyyy]")
ChronoAnnee = "[Chrono" & Format(Date, "yyyy]")
Me.Caption = Me.Caption & " Affaire = " & Affaire


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
sql = sql & "WHERE " & ChronoAnnee & ".[Clé ac] = " & Affaire & " and " & ChronoAnnee & ".[Clé ty] = 'PI' "
sql = sql & "ORDER BY " & ChronoAnnee & ".[Clé Ch] DESC;"
Set RsBaseNum = ConBaseNum.OpenRecordSet(sql)

While RsBaseNum.EOF = False
    txt5.AddItem "" & RsBaseNum![Clé ty] & "_" & RsBaseNum![Clé ac] & "_" & RsBaseNum![Année] & _
    "_" & RsBaseNum![Clé Ch] & "_" & RsBaseNum![Rév]
    
    txt5.List(txt5.ListCount - 1, 1) = "" & "" & RsBaseNum![Clé ty] & "_" & RsBaseNum![Clé ac] & "_" & RsBaseNum![Année] & _
    "_" & RsBaseNum![Clé Ch]
'    txt5.List(txt5.ListCount - 1, 1) = "" & RsBaseNum![Clé Ch]
    txt5.List(txt5.ListCount - 1, 2) = "" & RsBaseNum![Red_Nom] & " " & RsBaseNum![Red_P_Nom]
    txt5.List(txt5.ListCount - 1, 3) = "" & RsBaseNum![Verif_Nom] & " " & RsBaseNum![Verif_P_Nom]
    txt5.List(txt5.ListCount - 1, 4) = "" & RsBaseNum![Apr_Nom] & " " & RsBaseNum![Apr_P_Nom]
    txt5.List(txt5.ListCount - 1, 5) = "" & RsBaseNum![Rév]
     If strtxt7 = txt5.List(txt5.ListCount - 1, 0) Then txt5.ListIndex = txt5.ListCount - 1
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
sql = sql & "WHERE " & ChronoAnnee & ".[Clé ac] = " & Affaire & " and " & ChronoAnnee & ".[Clé ty] = 'PL' "
sql = sql & "ORDER BY " & ChronoAnnee & ".[Clé Ch] DESC;"
Set RsBaseNum = ConBaseNum.OpenRecordSet(sql)

While RsBaseNum.EOF = False
    txt6.AddItem "" & RsBaseNum![Clé ty] & "_" & RsBaseNum![Clé ac] & "_" & RsBaseNum![Année] & _
    "_" & RsBaseNum![Clé Ch] & "_" & RsBaseNum![Rév]
    
    
    txt6.List(txt6.ListCount - 1, 1) = "" & RsBaseNum![Clé ty] & "_" & RsBaseNum![Clé ac] & "_" & RsBaseNum![Année] & _
    "_" & RsBaseNum![Clé Ch]
    txt6.List(txt6.ListCount - 1, 2) = "" & RsBaseNum![Red_Nom] & " " & RsBaseNum![Red_P_Nom]
    txt6.List(txt6.ListCount - 1, 3) = "" & RsBaseNum![Verif_Nom] & " " & RsBaseNum![Verif_P_Nom]
    txt6.List(txt6.ListCount - 1, 4) = "" & RsBaseNum![Apr_Nom] & " " & RsBaseNum![Apr_P_Nom]
    txt6.List(txt6.ListCount - 1, 5) = " " & RsBaseNum![Rév]
    If strtxt8 = txt6.List(txt6.ListCount - 1, 0) Then txt6.ListIndex = txt6.ListCount - 1
    RsBaseNum.MoveNext
Wend
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
sql = sql & "WHERE " & ChronoAnnee & ".[Clé ac] = " & Affaire & " and " & ChronoAnnee & ".[Clé ty] = 'OU' "
sql = sql & "ORDER BY " & ChronoAnnee & ".[Clé Ch] DESC;"
Set RsBaseNum = ConBaseNum.OpenRecordSet(sql)

While RsBaseNum.EOF = False
    txt7.AddItem "" & RsBaseNum![Clé ty] & "_" & RsBaseNum![Clé ac] & "_" & RsBaseNum![Année] & _
    "_" & RsBaseNum![Clé Ch] & "_" & RsBaseNum![Rév]
     
    txt7.List(txt7.ListCount - 1, 1) = "" & RsBaseNum![Clé ty] & "_" & RsBaseNum![Clé ac] & "_" & RsBaseNum![Année] & _
    "_" & RsBaseNum![Clé Ch]
     txt7.List(txt7.ListCount - 1, 2) = " " & RsBaseNum![Rév]
     If strtxt9 = txt7.List(txt7.ListCount - 1, 0) Then txt7.ListIndex = txt7.ListCount - 1
    RsBaseNum.MoveNext
Wend

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
sql = sql & "WHERE " & ChronoAnnee & ".[Clé ac] = " & Affaire & " and " & ChronoAnnee & ".[Clé ty] = 'LI' "
sql = sql & "ORDER BY " & ChronoAnnee & ".[Clé Ch] DESC;"
Set RsBaseNum = ConBaseNum.OpenRecordSet(sql)

While RsBaseNum.EOF = False
    txt8.AddItem "" & RsBaseNum![Clé ty] & "_" & RsBaseNum![Clé ac] & "_" & RsBaseNum![Année] & _
    "_" & RsBaseNum![Clé Ch] & "_" & RsBaseNum![Rév]
   
    txt8.List(txt8.ListCount - 1, 1) = "" & RsBaseNum![Clé ty] & "_" & RsBaseNum![Clé ac] & "_" & RsBaseNum![Année] & _
    "_" & RsBaseNum![Clé Ch]
     txt8.List(txt8.ListCount - 1, 2) = " " & RsBaseNum![Rév]
       If strtxt10 = txt8.List(txt8.ListCount - 1, 0) Then txt8.ListIndex = txt8.ListCount - 1
    RsBaseNum.MoveNext
Wend

ConBaseNum.CloseConnection

Me.Show
Charge = boolExecute
End Function

Private Sub CommandButton2_Click()
Dim boolCahnge As Boolean
boolCahnge = False

If CopieStrtxt7 <> txt5 Then boolCahnge = True
If CopieStrtxt8 <> txt6 Then boolCahnge = True
If CopieStrtxt9 <> txt7 Then boolCahnge = True
If CopieStrtxt10 <> txt8 Then boolCahnge = True
If MyFormatQRY(DescIndice) = False Then Exit Sub
If boolCahnge = False Then
    MsgBox "Vous devez changer au moins un N° chrono dans une des liste", vbOKOnly + vbExclamation, "Erreur sur l'indice"
    Exit Sub
End If
boolExecute = True
Noquite = False
Noquite = False
Me.Hide
End Sub

Private Sub CommandButton3_Click()
Me.Hide
End Sub

Private Sub txt5_Click()
If Me.txt5.ListIndex <> -1 Then
    Me.txt10 = txt5.List(Me.txt5.ListIndex, 2)
     Me.txt11 = txt5.List(Me.txt5.ListIndex, 3)
      Me.txt12 = txt5.List(Me.txt5.ListIndex, 4)
Else
     Me.txt10 = ""
     Me.txt11 = ""
     Me.txt12 = ""
End If
End Sub


Private Sub UserForm_Activate()
Noquite = True
boolExecute = False
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Cancel = Noquite
End Sub

