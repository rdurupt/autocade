VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CeerXls 
   Caption         =   "Créer Nouvelle liste de Fils :"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4185
   OleObjectBlob   =   "CeerXls.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CeerXls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub CommandButton1_Click()

If Trim("" & Me.FichierXLS) = "" Then
    MsgBox "Vous devez saisir le chemin ainsi que le nom du fichier EXCEL à importer", vbExclamation, "Erreur"
    Me.FichierXLS.SetFocus
    Exit Sub
End If
'If Trim("" & Me.ProjetName) = "" Then
'    MsgBox "Vous devez saisir le nom du projet", vbExclamation, "Erreur"
'    Me.ProjetName.SetFocus
'    Exit Sub
'End If
'If Trim("" & Me.ProjetIndice) = "" Then
'    MsgBox "Vous devez saisir l'indice du projet", vbExclamation, "Erreur"
'    Me.ProjetIndice.SetFocus
'    Exit Sub
'End If
If UCase(Right(Trim("" & Me.FichierXLS), 4)) <> ".XLS" Then
     Me.FichierXLS = Trim("" & Me.FichierXLS) & ".XLS"
End If
Dim Fso As New FileSystemObject
    If Fso.FileExists(Trim("" & Me.FichierXLS)) = True Then
    
        If MsgBox(Me.FichierXLS & vbCrLf & "Existe déjà voulez vous le remplacer.", vbQuestion + vbYesNo, "Fichier Existe:") = vbNo Then
            Me.FichierXLS.SetFocus
            Set Fso = Nothing
            Exit Sub
        End If
    End If
Set Fso = Nothing
Dim Rs As Recordset
Dim Sql As String
'Con.OpenConnetion db
'Sql = "SELECT T_Projet.Projet, T_indiceProjet.Indice, T_indiceProjet.Approuver "
'Sql = Sql & "FROM T_Projet INNER JOIN T_indiceProjet ON T_Projet.id = T_indiceProjet.IdProjet "
'Sql = Sql & "WHERE T_Projet.Projet='" & Me.ProjetName & "' "
'Sql = Sql & "AND T_indiceProjet.Indice='" & Me.ProjetIndice & "'; "
'Set Rs = Con.OpenRecordSet(Sql)
'If Rs.EOF = False Then
'    If Rs!Approuver = True Then
'        MsgBox "Le projet : " & Me.ProjetName & " Indice : " & Me.ProjetIndice & " a déjà été approuver et ne peut pas être modifié.", vbCritical, "Import EXCEL"
'        ProjetIndice.SetFocus
'        Exit Sub
'    Else
'        If MsgBox("Le projet : " & Me.ProjetName & " Indice : " & Me.ProjetIndice & " existe déjà voulez vous le remplacer.", vbYesNo, "Import EXCEL") = vbNo Then
'            Set Rs = Con.CloseRecordSet(Rs)
'            Con.CloseConnection
'            Exit Sub
'        End If
'
'    End If
'End If
'Con.CloseConnection
If MsgBox("Voulez vous exécuter l'importation de :" & Me.FichierXLS, vbQuestion + vbYesNo, "Import EXCEL") = vbNo Then Exit Sub
Unload Me
DoEvents
UserForm2.Chargement Me.FichierXLS, True

End Sub

Private Sub CommandButton2_Click()
Unload Me
End Sub

