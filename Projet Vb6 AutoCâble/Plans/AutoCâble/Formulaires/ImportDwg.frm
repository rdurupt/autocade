VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ImportDwg 
   Caption         =   "Importer fichier AUTOCAD :"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8430
   OleObjectBlob   =   "ImportDwg.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ImportDwg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CleAc_Change()
Approbateur = Me.CleAc.List(CleAc.ListIndex, 1)
End Sub

Private Sub CommandButton1_Click()

If Trim("" & Me.FichierXLS) = "" Then
    MsgBox "Vous devez saisir le chemin ainsi que le nom du fichier AUTOCAD à importer", vbExclamation, "Erreur"
    Me.FichierXLS.SetFocus
    Exit Sub
End If
If Trim("" & Me.ProjetName) = "" Then
    MsgBox "Vous devez saisir le nom du projet", vbExclamation, "Erreur"
    Me.ProjetName.SetFocus
    Exit Sub
End If
'If Trim("" & Me.ProjetIndice) = "" Then
'    MsgBox "Vous devez saisir l'indice du projet", vbExclamation, "Erreur"
'    Me.ProjetIndice.SetFocus
'    Exit Sub
'End If
If Trim("" & txt20) = "" Then CommandButton3_Click
If Trim("" & txt20) = "" Then Exit Sub
Dim Fso As New FileSystemObject
If UCase(Right(Trim("" & Me.FichierXLS), 4)) <> ".DWG" Then
     Me.FichierXLS = Trim("" & Me.FichierXLS) & ".DWG"
End If
    If Fso.FileExists(Trim("" & Me.FichierXLS)) = False Then
      MsgBox "Le chemin ou le nom du fichier AUTOCAD introuvable." & vbCrLf & "Vérifiez l'orthographe ou l'existence de ce fichier", vbExclamation, "Erreur"
      Set Fso = Nothing
       Me.FichierXLS.SetFocus
      Exit Sub
    End If
Set Fso = Nothing
Dim Rs As Recordset
Dim Sql As String
Con.OpenConnetion db
Sql = "SELECT T_Projet.Projet, T_indiceProjet.Indice "
Sql = Sql & "FROM T_Projet INNER JOIN T_indiceProjet ON T_Projet.id = T_indiceProjet.IdProjet "
Sql = Sql & "WHERE T_Projet.Projet='" & Me.ProjetName & "' "
Sql = Sql & "AND T_indiceProjet.li='" & TextBox7 & CleAc & "_" & Annee & txt20 & "_" & REV & "_" & TextBox1 & "'; "
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = False Then
'    If Rs!Approuver = True Then
'        MsgBox "Le projet : " & Me.ProjetName & " Indice : " & Me.ProjetIndice & " a déjà été approuver et ne peut pas être modifié.", vbCritical, "Import AUTOCAD"
'        ProjetIndice.SetFocus
'        Exit Sub
'    Else
        If MsgBox("Le projet : " & MeTextBox1 & " Indice : " & MeTextBox1 & " existe déjà voulez vous le remplacer.", vbYesNo, "Import AUTOCAD") = vbNo Then
            Set Rs = Con.CloseRecordSet(Rs)
            Con.CloseConnection
            Exit Sub
        End If
        
'    End If
End If
Con.CloseConnection
If MsgBox("Voulez vous exécuter l'importation de :" & Me.FichierXLS, vbQuestion + vbYesNo, "Import AUTOCAD") = vbNo Then Exit Sub
Unload Me
DoEvents
ScanDessin Me.FichierXLS, Me.ProjetName, Me.TextBox1, Me.ProjetDesciption, TextBox7 & CleAc & "_" & Annee & txt20 & "_" & REV & "_" & TextBox1, CleAc

End Sub

Private Sub CommandButton2_Click()
Unload Me
End Sub

Private Sub CommandButton3_Click()
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
Sql = Sql & "Année, Rév, [Clé ty],  NumAuToCad, "
Sql = Sql & "[Clé ac] ,[Objet],Destinataire,rv )"
Sql = Sql & "values( '" & Format(Date, "yy") & "', '" & MyReplace(Me.TextBox1) & " ', 'LI'," & Rs!NumErreur & ","
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

Private Sub UserForm_Activate()
Dim Sql As String
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

