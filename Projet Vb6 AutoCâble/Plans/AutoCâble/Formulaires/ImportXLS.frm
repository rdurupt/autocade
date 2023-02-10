VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ImportXLS 
   Caption         =   "Créer un plan :"
   ClientHeight    =   9360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8025
   OleObjectBlob   =   "ImportXLS.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ImportXLS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim boolChrono As Boolean

Private Sub CleAc_Change()
Approbateur = Me.CleAc.List(CleAc.ListIndex, 1)
End Sub

Private Sub CommandButton1_Click()
Dim TxtOption As String
Con.OpenConnetion db
 Set TableauPath = funPath
 Con.CloseConnection
If Me.OptionButton1.Value = True Then
Me.Hide
    CartoucheEncelade.Show
GoTo Fin

    TxtOption = "A"
End If
If ValideChampsTexte(Me, 2) = False Then Exit Sub

If Me.OptionButton2.Value = True Then
    TxtOption = "E"
End If
If Me.OptionButton3.Value = True Then
    TxtOption = "N"
End If
If Trim("" & txt20) = "" Then CommandButton3_Click
             If Trim("" & txt20) = "" Then Exit Sub
If Trim("" & Me.FichierXLS) = "" Then
    MsgBox "Vous devez saisir le chemin ainsi que le nom du fichier EXCEL à importer", vbExclamation, "Erreur"
    Me.FichierXLS.SetFocus
    Exit Sub
End If

If Trim("" & Me.ProjetName) = "" Then
    MsgBox "Vous devez saisir le nom du projet", vbExclamation, "Erreur"
    Me.ProjetName.SetFocus
    Exit Sub
End If

If Trim("" & Me.ProjetIndice) = "" Then
    MsgBox "Vous devez saisir l'indice du projet", vbExclamation, "Erreur"
    Me.ProjetIndice.SetFocus
    Exit Sub
End If

If UCase(Right(Trim("" & Me.FichierXLS), 4)) <> ".XLS" Then
     Me.FichierXLS = Trim("" & Me.FichierXLS) & ".XLS"
End If

Dim Fso As New FileSystemObject
    If Fso.FileExists(Trim("" & Me.FichierXLS)) = False Then
        If TxtOption = "E" Then
      MsgBox "Le chemin ou le nom du fichier EXCEL introuvable." & vbCrLf & "Vérifiez l'orthographe ou l'existence de ce fichier", vbExclamation, "Erreur"
      Set Fso = Nothing
       Me.FichierXLS.SetFocus
      Exit Sub
        End If
      
      Else
        If TxtOption = "N" Then
            If MsgBox(Me.FichierXLS & vbCrLf & "Existe déjà voulez vous le remplacer.", vbQuestion + vbYesNo, "Fichier Existe:") = vbNo Then
                Me.FichierXLS.SetFocus
                Set Fso = Nothing
                Exit Sub
            End If
        End If
    End If
    
Set Fso = Nothing
Dim Rs As Recordset
Dim Sql As String
Con.OpenConnetion db
'Sql = "SELECT T_Projet.Projet, T_indiceProjet.Indice "
'Sql = Sql & "FROM T_Projet INNER JOIN T_indiceProjet ON T_Projet.id = T_indiceProjet.IdProjet "
'Sql = Sql & "WHERE T_Projet.Projet='" & Me.ProjetName & "' "
'Sql = Sql & "AND T_indiceProjet.Indice='" & Me.ProjetIndice & "'; "
'Set Rs = Con.OpenRecordSet(Sql)
'If Rs.EOF = False Then
''    If Rs!Approuver = True Then
''        MsgBox "Le projet : " & Me.ProjetName & " Indice : " & Me.ProjetIndice & " a déjà été approuver et ne peut pas être modifié.", vbCritical, "Import EXCEL"
''        ProjetIndice.SetFocus
''        Exit Sub
''    Else
'        If MsgBox("Le projet : " & Me.ProjetName & " Indice : " & Me.ProjetIndice & " existe déjà voulez vous le remplacer.", vbYesNo, "Import EXCEL") = vbNo Then
'            Set Rs = Con.CloseRecordSet(Rs)
'            Con.CloseConnection
'            Exit Sub
'        End If
'
''    End If
'End If
'
'Con.CloseConnection
If MsgBox("Voulez vous exécuter l'importation de :" & Me.FichierXLS, vbQuestion + vbYesNo, "Import EXCEL") = vbNo Then Exit Sub

DoEvents
Select Case TxtOption
         Case "A"
         Case "E"
                 Me.Hide
                 ImporteXls Me.FichierXLS, Me.ProjetName, Me.ProjetIndice, Me.ProjetDesciption, TextBox7 & CleAc & "_" & Annee & txt20 & "_" & REV, CleAc.Text
          
         
         Case "N"
            UserForm2.Chargement Me.FichierXLS, Me.ComboBox1.Text, True
         If UserForm2.boolExcute = True Then
          Unload UserForm2
          Me.Hide
                ImporteXls Me.FichierXLS, Me.ProjetName, Me.ProjetIndice, Me.ProjetDesciption, TextBox7 & CleAc & "_" & Annee & txt20 & "_" & REV, CleAc.Text
          Else
             Unload UserForm2
             Exit Sub
          
          End If
End Select

CartoucheEncelade.Charge Me.ProjetName, TextBox7 & CleAc & "_" & Annee & txt20 & "_" & REV, Me.txt0, Me.txt1, Me.txt2, True

Fin:
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
    If Trim("" & CleAc) = "" Then
            MsgBox "Valeur de : Avtivité obligatoire", vbExclamation
            Me.CleAc.SetFocus
          Exit Sub
        End If
    If Trim("" & Approbateur) = "" Then
            MsgBox "Valeur de : Approbateur obligatoire", vbExclamation
            Me.Approbateur.SetFocus
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

Private Sub OptionButton1_Click()
    EXCEL.Enabled = False
End Sub

Private Sub OptionButton2_Click()
EXCEL.Enabled = True
End Sub

Private Sub OptionButton3_Click()
EXCEL.Enabled = True
End Sub

Private Sub UserForm_Activate()
Dim Sql As String
OptionButton1.Value = True
'OptionButton1_Click
Dim RsBaseNum As Recordset
ConNumPlan.OpenConnetion DbNumPlan
Sql = "SELECT Activité.[Clé ac], Responsables.Nom, Responsables.Prénom "
Sql = Sql & "FROM Responsables RIGHT JOIN Activité ON Responsables.[Clé res] = Activité.[Clé re] "
Sql = Sql & "WHERE Activité.[Lib ac] Is Not Null "
Sql = Sql & "And Activité.[St ac] <> 4 "
Sql = Sql & "ORDER BY Activité.[Clé ac] DESC;"
Set RsBaseNum = ConNumPlan.OpenRecordSet(Sql)
CleAc.AddItem ""
CleAc.List(CleAc.ListCount - 1, 1) = "0"
While RsBaseNum.EOF = False
    CleAc.AddItem CStr(RsBaseNum![Clé ac])
    CleAc.List(CleAc.ListCount - 1, 1) = Trim("" & RsBaseNum!Nom) & " " & Trim("" & RsBaseNum!Prénom)
     If Me.CleAc.ListCount = 1 Then Me.CleAc.Text = CStr(RsBaseNum![Clé ac])
    RsBaseNum.MoveNext
Wend
Me.Annee = Format(Date, "yy_")
Set RsBaseNum = ConNumPlan.CloseRecordSet(RsBaseNum)
 Sql = "SELECT T_Clients.Client, T_Clients.Formulaire "
Sql = Sql & "FROM T_Clients "
Sql = Sql & "ORDER BY T_Clients.Client;"
Con.OpenConnetion db
Set RsBaseNum = Con.OpenRecordSet(Sql)
While RsBaseNum.EOF = False
    Me.ComboBox1.AddItem Trim("" & RsBaseNum!Client)
        If Me.ComboBox1.ListCount = 1 Then Me.ComboBox1.Text = Trim("" & RsBaseNum!Client)

    RsBaseNum.MoveNext
Wend
Set RsBaseNum = Con.CloseRecordSet(RsBaseNum)
ConNumPlan.CloseConnection
Con.CloseConnection
End Sub

