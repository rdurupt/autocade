VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EDITER 
   Caption         =   "Editer un plan :"
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7980
   OleObjectBlob   =   "EDITER.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "EDITER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub CheckBox1_Click()
If Me.CheckBox1.Value = True Then
    Me.Frame1.Enabled = False
Else
    Me.Frame1.Enabled = True
End If
End Sub

Private Sub CleAc_Change()
Approbateur = Me.CleAc.List(CleAc.ListIndex, 1)

End Sub

Private Sub CommandButton1_Click()
Dim Fso As New FileSystemObject
Dim pathTmpXls As String
If lstIndice.Text = "" Then Exit Sub
 If Me.Caption = "Modifier un plan :" Then

    If txt20 = "" Then CommandButton4_Click
    If txt20 = "" Then Exit Sub
        pathTmpXls = Environ("USERPROFILE") & "\Mes Documents\" & TextBox7 & CleAc & "_" & Annee & txt20 & "_" & REV & ".XLS"
    Else
        pathTmpXls = Environ("USERPROFILE") & "\Mes Documents\" & lstIndice & ".XLS"

 End If
 
If Fso.FileExists(pathTmpXls) = True Then
    Fso.DeleteFile pathTmpXls
End If
varProjet = Me.lstProjets
If lstIndice.ListIndex < 0 Then Exit Sub
varIndice = Me.lstIndice.List(lstIndice.ListIndex, 0)
Me.Hide
 ExporteXls pathTmpXls, varProjet, varIndice
 UserForm2.Chargement pathTmpXls, Me.ComboBox1.Text
 If UserForm2.boolExcute = False Then
   Unload UserForm2
    Exit Sub
 End If
 
 If Me.Caption = "Modifier un plan :" Then
 If Trim("" & ProjetName) <> "" Then
        ImporteXls pathTmpXls, Me.ProjetName, Me.ProjetIndice, Me.ProjetDesciption, TextBox7 & CleAc & "_" & Annee & txt20 & "_" & REV, CleAc.Text
        CartoucheEncelade.Charge Me.ProjetName, TextBox7 & CleAc & "_" & Annee & txt20 & "_" & REV, Me.txt0, Me.txt1, Me.txt2, True
Else
   a = Me.lstIndice
            ImporteXls pathTmpXls, Me.lstProjets, Me.ProjetIndice, Me.ProjetDesciption, TextBox7 & CleAc & "_" & Annee & txt20 & "_" & REV, CleAc.Text
        CartoucheEncelade.Charge Me.lstProjets, TextBox7 & CleAc & "_" & Annee & txt20 & "_" & REV, Me.txt0, Me.txt1, Me.txt2, True

End If
    Else
   
        ImporteXls pathTmpXls, Me.lstProjets, Me.lstIndice.List(lstIndice.ListIndex, 1), Me.lstProjets.List(Me.lstProjets.ListIndex, 1), Me.lstIndice, Me.lstProjets.List(Me.lstProjets.ListIndex, 2)
        CartoucheEncelade.Charge Me.lstProjets, Me.lstIndice.List(lstIndice.ListIndex, 0), Me.txt0, Me.txt1, Me.txt2, True

    End If
    Unload UserForm2
End Sub


Private Sub CommandButton3_Click()
Me.Hide
End Sub

Private Sub Label1_Click()

End Sub

Private Sub CommandButton4_Click()
If ValideChampsTexte(Me, 2) = False Then Exit Sub
Dim Rs As Recordset
Dim RsNum As Recordset
Dim Sql As String
Con.OpenConnetion db
ConNumPlan.OpenConnetion DbNumPlan
  If boolChrono = True Then Exit Sub
    If Trim("" & Approbateur) = "" Then
            MsgBox "Valeur de : Approbateur obligatoire", vbExclamation
            Approbateur.SetFocus
          Exit Sub
        End If
       boolChrono = True
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
Sql = Sql & "Société,Année, Rév, [Clé ty],  NumAuToCad, "
Sql = Sql & "[Clé ac] ,[Objet],Destinataire,rv )"
Sql = Sql & "values('" & MyReplace(ComboBox1) & "', '" & Format(Date, "yy") & "', '" & MyReplace(Me.ProjetIndice) & " ', 'LI'," & Rs!NumErreur & ","
Sql = Sql & Me.CleAc.Text & ", ' Vague: " & MyReplace(txt0) & " Ensemble: " & MyReplace(txt1) & "  Equipement: " & MyReplace(txt2) & "','" & MyReplace(Approbateur) & "','" & MyReplace(REV) & "');"
Debug.Print Sql
ConNumPlan.Exequte Sql
Sql = "SELECT Chrono" & Format(Date, "yyyy") & ".[Clé Ch] FROM Chrono" & Format(Date, "yyyy") & " WHERE Chrono" & Format(Date, "yyyy") & ".NumAuToCad= " & Rs!NumErreur & ";"
Set Rs = ConNumPlan.OpenRecordSet(Sql)
If Rs.EOF = False Then
    txt20 = Rs![Clé Ch]
End If
Set Rs = ConNumPlan.CloseRecordSet(Rs)
End If
Con.CloseConnection
ConNumPlan.CloseConnection
End Sub

Private Sub CommandButton5_Click()
Con.OpenConnetion db
UserForm1.Charger txt2, ";", "Equipement:", " "
Con.CloseConnection
End Sub

Private Sub CommandButton6_Click()
Con.OpenConnetion db
UserForm1.Charger txt1, vbCrLf, "Ensemble:"
Con.CloseConnection
End Sub

Private Sub CommandButton7_Click()
Con.OpenConnetion db
UserForm1.Charger txt0, " ", "Vagues:", " "

Con.CloseConnection

End Sub

Private Sub CommandButton8_Click()
UserForm3.Show
Maj Me.ComboBox1

End Sub
Sub Maj(MyControl As Object)
Dim Rs As Recordset
Dim Sql As String
Con.OpenConnetion db
MyControl.Clear
Sql = "SELECT T_Clients.Client, T_Clients.Formulaire "
Sql = Sql & "FROM T_Clients "
Sql = Sql & "ORDER BY T_Clients.Client;"

Set Rs = Con.OpenRecordSet(Sql)
While Rs.EOF = False
    MyControl.AddItem Trim("" & Rs!Client)
        If MyControl.ListCount = 1 Then MyControl.Text = Trim("" & Rs!Client)

    Rs.MoveNext
Wend
Set Rs = Con.CloseRecordSet(Rs)
Con.CloseConnection
End Sub

Private Sub Frame1_Click()

End Sub

Private Sub lstProjets_Click()
Dim Rs As Recordset
Dim Sql As String
Dim indexClient As Long
Sql = "SELECT T_indiceProjet.Li ,T_indiceProjet.Indice "
Sql = Sql & "FROM T_Projet INNER JOIN T_indiceProjet ON T_Projet.id = T_indiceProjet.IdProjet "
Sql = Sql & "WHERE T_Projet.Projet = '" & Me.lstProjets.Text & "' AND T_indiceProjet.IdStatus": If Me.Caption = "Modifier un plan :" Then Sql = Sql & "=3 " Else Sql = Sql & "<3 "
Sql = Sql & "ORDER BY T_indiceProjet.Indice;"


Set Rs = Con.OpenRecordSet(Sql)
Me.lstIndice.Clear
While Rs.EOF = False

    Me.lstIndice.AddItem Trim("" & Rs!LI)
    Me.lstIndice.List(Me.lstIndice.ListCount - 1, 1) = Trim("" & Rs!Indice)
    If Me.lstIndice.ListCount = 1 Then Me.lstIndice.ListIndex = 0
    Rs.MoveNext
Wend

End Sub




Private Sub UserForm_Activate()
Dim Rs As Recordset
Dim Sql As String
Dim indexClient As Long
Con.OpenConnetion db
 If Me.Caption = "Modifier un plan :" Then
    Frame1.Enabled = True
    Else
         Frame1.Enabled = False
 End If
 
 Sql = "SELECT T_Clients.Client, T_Clients.Formulaire "
Sql = Sql & "FROM T_Clients "
Sql = Sql & "ORDER BY T_Clients.Client;"

Set Rs = Con.OpenRecordSet(Sql)
While Rs.EOF = False
    Me.ComboBox1.AddItem Trim("" & Rs!Client)
        If Me.ComboBox1.ListCount = 1 Then Me.ComboBox1.Text = Trim("" & Rs!Client)

    Rs.MoveNext
Wend
Set Rs = Con.CloseRecordSet(Rs)
Sql = "SELECT T_Projet.id, T_Projet.Projet, T_Projet.Description ,T_Projet.CleAc FROM T_Projet "
Sql = Sql & "ORDER BY  T_Projet.Projet;"

Set Rs = Con.OpenRecordSet(Sql)
While Rs.EOF = False

    Me.lstProjets.AddItem Trim("" & Rs!Projet)
    Me.lstProjets.List(Me.lstProjets.ListCount - 1, 1) = Trim("" & Rs!Description)
    Me.lstProjets.List(Me.lstProjets.ListCount - 1, 2) = Trim("" & Rs!CleAc)
    If Me.lstProjets.ListCount = 1 Then Me.lstProjets.ListIndex = 0
    Rs.MoveNext
Wend

'OptionButton1.Value = True
'OptionButton1_Click
Dim RsBaseNum As Recordset
ConNumPlan.OpenConnetion DbNumPlan
Sql = "SELECT Activité.[Clé ac], Responsables.Nom, Responsables.Prénom "
Sql = Sql & "FROM Responsables RIGHT JOIN Activité ON Responsables.[Clé res] = Activité.[Clé re] "
Sql = Sql & "WHERE Activité.[Lib ac] Is Not Null "
Sql = Sql & "And Activité.[St ac] <> 4 "
Sql = Sql & "ORDER BY Activité.[Clé ac] DESC;"
Set RsBaseNum = ConNumPlan.OpenRecordSet(Sql)
While RsBaseNum.EOF = False
    CleAc.AddItem CStr(RsBaseNum![Clé ac])
    CleAc.List(CleAc.ListCount - 1, 1) = Trim("" & RsBaseNum!Nom) & " " & Trim("" & RsBaseNum!Prénom)
     If Me.CleAc.ListCount = 1 Then Me.CleAc.Text = CStr(RsBaseNum![Clé ac])
    RsBaseNum.MoveNext
Wend
Me.Annee = Format(Date, "yy_")
Set RsBaseNum = ConNumPlan.CloseRecordSet(RsBaseNum)
ConNumPlan.CloseConnection
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Con.CloseConnection
End Sub
Public Sub Charger(MeCaption As String)
Me.Caption = MeCaption
Me.Show
End Sub
