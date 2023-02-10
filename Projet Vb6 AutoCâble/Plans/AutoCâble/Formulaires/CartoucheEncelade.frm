VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CartoucheEncelade 
   Caption         =   "CARTOUCHE ENCELADE :"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10350
   OleObjectBlob   =   "CartoucheEncelade.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "CartoucheEncelade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public NumAuToCad As Long
Dim MyProjet As String
Dim MyLi As String

Private Sub CleAc_Change()
CleAc_Click
End Sub
Private Sub CleAc_Click()
If Me.CleAc.ListIndex <> -1 Then
    Me.txt2 = CleAc.List(Me.CleAc.ListIndex, 1)
Else
     Me.txt2 = ""
End If
txt12.ListIndex = 0
For I = 0 To txt12.ListCount - 1
    If UCase(txt12.List(I, 3) & " " & txt12.List(I, 1)) = UCase(CleAc.List(Me.CleAc.ListIndex, 2)) Then
        txt12.ListIndex = I
        Exit For
    End If
Next I
End Sub

Private Sub CombLi_Click()
If CombLi.ListCount > 0 Then
    Me.txt13 = CombLi.List(CombLi.ListIndex, 2)
    txt16 = CombLi.List(CombLi.ListIndex, 2)
End If
End Sub

Private Sub CommandButton1_Click()
Dim Sql As String
'If Trim("" & Me.txt20) = "" Then CommandButton2_Click
'If Trim("" & Me.txt20) = "" Then
'    Exit Sub
'Else
'    Me.txt15 = TextBox7 & CleAc.Text & "_" & Annee & txt20
'End If
If ValideChampsTexte(Me, 18) = False Then Exit Sub
LeCartoucheE = "CARTOUCHE ENCELADE.dwg"
Dim MyTag
 
NoClose = False
NbContolClient = 18
'sql = "UPDATE Chrono" & Format(Date, "yyyy") & " "
'sql = sql & "SET Chrono" & Format(Date, "yyyy") & ".Rév = '"
'sql = sql & MyReplace(Me.txt16) & " ', Chrono" & Format(Date, "yyyy") & ".Société = '"
'sql = sql & MyReplace(Me.txt1) & "', Chrono" & Format(Date, "yyyy") & ".Destinataire = '"
'sql = sql & MyReplace(Me.txt2) & "', Chrono" & Format(Date, "yyyy") & ".[Clé ac] = "
'sql = sql & Me.CleAc & ", Chrono" & Format(Date, "yyyy") & ".[Clé re] = "
'sql = sql & Me.txt8.Column(2) & ", Chrono" & Format(Date, "yyyy") & ".[Clé ve] = "
'sql = sql & Me.Txt10.Column(2) & ", Chrono" & Format(Date, "yyyy") & ".[Clé ap] = "
'sql = sql & Me.txt12.Column(2) & ", Chrono" & Format(Date, "yyyy") & ".Objet = ' Vague: " & MyReplace(txt3) & " Ensemble: " & MyReplace(txt14) & "  Equipement: " & MyReplace(txt6) & "' "
'sql = sql & "WHERE Chrono" & Format(Date, "yyyy") & ".[Clé Ch]=" & Me.txt20 & ";"
'ConNumPlan.Exequte sql


varProjet = Me.txt4.List(Me.txt4.ListIndex, 0)
varIndice = Me.txt13
  Con.CloseConnection
   ConNumPlan.CloseConnection
    Select Case UCase(Me.txt1.Text)
        Case "RENAULT"
            LeCient = UCase(Me.txt1.Text)
              LeCartouche = "CARTOUCHE  RENAULT.dwg"
            Set MyCARTOUCHE_Client = New CARTOUCHE_RENAULT
            boolFormClient = True
        Case Else
             LeCient = "RENAULT"
            boolFormClient = False
    End Select
   If boolFormClient = True Then
    Load MyCARTOUCHE_Client
    MyCARTOUCHE_Client.Show
    End If
    If VarPreced = True Then
        Set MyCARTOUCHE_Client = Nothing
        Exit Sub

    End If
    ConNumPlan.OpenConnetion DbNumPlan
   Con.OpenConnetion db
    Me.Hide
    subDessinerPlan

End Sub


Private Sub CommandButton2_Click()
Dim Rs As Recordset
Dim RsNum As Recordset
Dim Sql As String

       MyTag = Split(Me.txt2.tag, ";")

    If Trim("" & txt2) = "" Then
            MsgBox "Valeur de : " & MyTag(1) & " obligatoire", vbExclamation
            txt2.SetFocus
          Exit Sub
        End If
If NumAuToCad = 0 Then
Sql = "SELECT  T_NumErreur.NumErreur FROM T_NumErreur "
Sql = Sql & "WHERE T_NumErreur.LibErreur='NumAuto';"
Set Rs = Con.OpenRecordSet(Sql)
Sql = "UPDATE T_NumErreur SET T_NumErreur.NumErreur = [NumErreur]+1 "
Sql = Sql & "WHERE T_NumErreur.LibErreur='NumAuto';"
Con.Exequte Sql
Rs.Requery
NumAuToCad = Rs!NumErreur

Sql = "INSERT INTO Chrono" & Format(Date, "yyyy") & " ( "
Sql = Sql & "Année, Rév, [Clé ty], Société, Destinataire, NumAuToCad, "
Sql = Sql & "[Clé ac], [Clé re], [Clé ve], [Clé ap] )"
Sql = Sql & "values( '" & Format(Date, "yy") & "', '" & MyReplace(Me.txt16) & " ', 'PL','" & MyReplace(Me.txt1) & "','" & MyReplace(Me.txt2) & "', " & Rs!NumErreur & ","
Sql = Sql & Me.CleAc & ", " & Me.txt8.Column(2) & ", " & Me.Txt10.Column(2) & ", " & Me.txt12.Column(2) & ");"
Debug.Print Sql
ConNumPlan.Exequte Sql
Sql = "SELECT Chrono" & Format(Date, "yyyy") & ".[Clé Ch] FROM Chrono" & Format(Date, "yyyy") & " WHERE Chrono" & Format(Date, "yyyy") & ".NumAuToCad= " & Rs!NumErreur & ";"
Set Rs = ConNumPlan.OpenRecordSet(Sql)
If Rs.EOF = False Then
    txt20 = Rs![Clé Ch]
End If
Set Rs = ConNumPlan.CloseRecordSet(Rs)
End If
End Sub

Private Sub CommandButton3_Click()
UserForm1.Charger txt14, vbCrLf, "Ensemble:"
End Sub

Private Sub CommandButton4_Click()
UserForm1.Charger txt6, ";", "Equipement:", "_"
End Sub



Private Sub CommandButton5_Click()
UserForm1.Charger txt3, " ", "Vagues:", " "
End Sub

Private Sub CommandButton6_Click()
UserForm3.Show
Maj Me.txt1
End Sub



Private Sub txt10_Click()
If Trim("" & Me.Txt10.Text) = "" Then
     Me.txt9 = ""
Else
    Me.txt9 = Date
End If

End Sub


Private Sub txt12_Click()
If Trim("" & Me.txt12.Text) = "" Then
     Me.txt11 = ""
Else
    Me.txt11 = Date
End If

End Sub


Private Sub txt13_Click()
txt16 = txt13.Text
End Sub





Private Sub txt15_Change()

End Sub

Private Sub txt4_Click()
Dim Rs As Recordset
Dim Sql As String
Dim indexClient As Long
Sql = "SELECT T_indiceProjet.Indice,T_indiceProjet.Li, T_Status.Status "
Sql = Sql & "FROM T_Status INNER JOIN (T_Projet INNER JOIN T_indiceProjet ON T_Projet.id = T_indiceProjet.IdProjet) ON T_Status.Id = T_indiceProjet.IdStatus "
Sql = Sql & "WHERE T_Projet.CleAc = " & Me.txt4.List(Me.txt4.ListIndex, 2) & " and T_Projet.Id=  " & Me.txt4.List(Me.txt4.ListIndex, 1) & " "
Sql = Sql & "ORDER BY T_indiceProjet.Li;"


Set Rs = Con.OpenRecordSet(Sql)
Me.CombLi.Clear
While Rs.EOF = False

    Me.CombLi.AddItem Trim("" & Rs!Status) & " : " & Trim("" & Rs!LI)
    Me.CombLi.List(Me.CombLi.ListCount - 1, 1) = Trim("" & Rs!LI)
     Me.CombLi.List(Me.CombLi.ListCount - 1, 2) = Trim("" & Rs!Indice)
     
    If Me.CombLi.ListCount = 1 Then
        Me.CombLi.ListIndex = 0
    End If
    Rs.MoveNext
Wend
Set Rs = Con.CloseRecordSet(Rs)
DoEvents
For I = 0 To CleAc.ListCount - 1
    If Me.txt4.List(Me.txt4.ListIndex, 2) = CleAc.List(I) Then
        CleAc.ListIndex = I
    End If
Next I
End Sub



Private Sub txt8_Click()
If Trim("" & Me.txt8.Text) = "" Then
     Me.txt7 = ""
Else
    Me.txt7 = Date
End If

End Sub

Private Sub UserForm_Activate()
Dim Rs As Recordset
Dim RsBaseNum As Recordset
Dim Sql As String
Dim indexClient As Long
'If boolCreationPlan = False Then
    Me.CombLi.Locked = boolCreationPlan
'Else
'     Me.txt13.Locked = True
'End If
Con.OpenConnetion db
ConNumPlan.OpenConnetion DbNumPlan
Sql = "SELECT Agent.[Clé ag], Agent.[Nom ag], Agent.[Prénom ag] "
Sql = Sql & "FROM Agent "
Sql = Sql & "ORDER BY Agent.[Nom ag];"
Set RsBaseNum = ConNumPlan.OpenRecordSet(Sql)
    Me.txt8.AddItem " "
    Me.Txt10.AddItem " "
    Me.txt12.AddItem " "
    Me.txt8.List(Me.txt8.ListCount - 1, 1) = " "
     Me.txt8.ListIndex = 0
      Me.Txt10.ListIndex = 0
       Me.txt12.ListIndex = 0
    Me.txt8.List(Me.txt8.ListCount - 1, 2) = "0"
    Me.Txt10.List(Me.Txt10.ListCount - 1, 1) = " "
    Me.Txt10.List(Me.Txt10.ListCount - 1, 2) = "0"
     Me.txt12.List(Me.txt12.ListCount - 1, 1) = " "
    Me.txt12.List(Me.txt12.ListCount - 1, 2) = "0"
While RsBaseNum.EOF = False
   
     If Trim("" & RsBaseNum![Prénom ag]) = "" Then
        If Trim("" & RsBaseNum![nom ag]) = "" Then
            Me.txt8.AddItem ""
            Me.Txt10.AddItem ""
            Me.txt12.AddItem ""
        Else
            If UCase("A Recruter") = UCase(Trim("" & RsBaseNum![nom ag])) Then
                 Me.txt8.AddItem Trim("" & RsBaseNum![nom ag])
                Me.Txt10.AddItem Trim("" & RsBaseNum![nom ag])
                Me.txt12.AddItem Trim("" & RsBaseNum![nom ag])
            Else
                 Me.txt8.AddItem "?." & Trim("" & RsBaseNum![nom ag])
                Me.Txt10.AddItem "?." & Trim("" & RsBaseNum![nom ag])
                Me.txt12.AddItem "?." & Trim("" & RsBaseNum![nom ag])
            End If
        End If
     Else
         If Trim("" & RsBaseNum![nom ag]) = "" Then
            Me.txt8.AddItem Trim("" & RsBaseNum![Prénom ag]) & ".?"
            Me.Txt10.AddItem Trim("" & RsBaseNum![Prénom ag]) & ".?"
            Me.txt12.AddItem Trim("" & RsBaseNum![Prénom ag]) & ".?"

           
        Else
         Me.txt8.AddItem UCase(Left(Trim("" & RsBaseNum![Prénom ag]), 1)) & "." & Trim("" & RsBaseNum![nom ag])
         Me.Txt10.AddItem UCase(Left(Trim("" & RsBaseNum![Prénom ag]), 1)) & "." & Trim("" & RsBaseNum![nom ag])
         Me.txt12.AddItem UCase(Left(Trim("" & RsBaseNum![Prénom ag]), 1)) & "." & Trim("" & RsBaseNum![nom ag])

        End If
     End If
    Me.txt8.List(Me.txt8.ListCount - 1, 1) = Trim("" & RsBaseNum![Prénom ag])
    Me.txt8.List(Me.txt8.ListCount - 1, 2) = Trim("" & RsBaseNum![Clé ag])
    Me.txt8.List(Me.txt8.ListCount - 1, 3) = Trim("" & RsBaseNum![nom ag])
    Me.Txt10.List(Me.Txt10.ListCount - 1, 1) = Trim("" & RsBaseNum![Prénom ag])
    Me.Txt10.List(Me.Txt10.ListCount - 1, 2) = Trim("" & RsBaseNum![Clé ag])
    Me.Txt10.List(Me.Txt10.ListCount - 1, 3) = Trim("" & RsBaseNum![nom ag])
    Me.txt12.List(Me.txt12.ListCount - 1, 1) = Trim("" & RsBaseNum![Prénom ag])
    Me.txt12.List(Me.txt12.ListCount - 1, 2) = Trim("" & RsBaseNum![Clé ag])
    Me.txt12.List(Me.txt12.ListCount - 1, 3) = Trim("" & RsBaseNum![nom ag])

If Me.txt12.ListCount = 1 Then
    Me.txt8.ListIndex = 0
    Me.Txt10.ListIndex = 0
    Me.txt12.ListIndex = 0
End If
    RsBaseNum.MoveNext

Wend

indexClient = 0
Sql = "SELECT Activité.[Clé ac], Responsables.Nom, Responsables.Prénom ,Activité.Int "
Sql = Sql & "FROM Responsables RIGHT JOIN Activité ON Responsables.[Clé res] = Activité.[Clé re] "
Sql = Sql & "WHERE Activité.[Lib ac] Is Not Null "
Sql = Sql & "And Activité.[St ac] <> 4 "
Sql = Sql & "ORDER BY Activité.[Clé ac] DESC;"
Set RsBaseNum = ConNumPlan.OpenRecordSet(Sql)
While RsBaseNum.EOF = False
    CleAc.AddItem CStr(RsBaseNum![Clé ac])
    CleAc.List(CleAc.ListCount - 1, 1) = Trim("" & RsBaseNum!Int)
    CleAc.List(CleAc.ListCount - 1, 2) = Trim("" & RsBaseNum!Nom) & " " & Trim("" & RsBaseNum!Prénom)
     If Me.CleAc.ListCount = 1 Then Me.CleAc.Text = CStr(RsBaseNum![Clé ac])
    RsBaseNum.MoveNext
Wend

Sql = "SELECT T_Projet.id, T_Projet.Projet, T_Projet.Description , T_Projet.CleAc FROM T_Projet "
Sql = Sql & "ORDER BY  T_Projet.Projet;"

Set Rs = Con.OpenRecordSet(Sql)
While Rs.EOF = False

    Me.txt4.AddItem Trim("" & Rs!Projet)
     Me.txt4.List(Me.txt4.ListCount - 1, 1) = Trim("" & Rs!Id)
    Me.txt4.List(Me.txt4.ListCount - 1, 2) = Trim("" & Rs!CleAc)
    If Me.txt4.ListCount = 1 Then Me.txt4.Text = Trim("" & Rs!Projet)
    Rs.MoveNext
Wend
Sql = "SELECT T_Clients.Client, T_Clients.Formulaire "
Sql = Sql & "FROM T_Clients "
Sql = Sql & "ORDER BY T_Clients.Client;"

Set Rs = Con.OpenRecordSet(Sql)
 Me.txt1.AddItem ""
While Rs.EOF = False
    Me.txt1.AddItem Trim("" & Rs!Client)
        If Me.txt1.ListCount = 1 Then Me.txt1.Text = Trim("" & Rs!Client)

    Rs.MoveNext
Wend
Set Rs = Con.CloseRecordSet(Rs)


Annee = Format(Date, "yy") & "_"
'ComboBox1
Set RsBaseNum = ConNumPlan.CloseRecordSet(RsBaseNum)
For I = 0 To Me.txt4.ListCount - 1
If Me.txt4.List(I) = MyProjet Then Me.txt4.ListIndex = I
Next I
For I = 0 To Me.CombLi.ListCount - 1
If Me.CombLi.List(I) = MyLi Then Me.CombLi.ListIndex = I
Next I


End Sub

Public Sub Charge(Optional Projet As String, Optional LI As String, Optional txt0 As String, Optional txt1 As String, Optional txt2 As String, Optional Locked As Boolean)
MyProjet = Projet
MyLi = LI
boolCreationPlan = Locked

Me.txt3 = txt0
Me.txt14 = txt1
Me.txt6 = txt2
Me.Show
End Sub
