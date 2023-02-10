VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExportXls 
   Caption         =   "Exporter vers fichier EXCEL :"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3015
   OleObjectBlob   =   "ExportXls.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "ExportXls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub CommandButton1_Click()
Dim Fso As New FileSystemObject
If Trim("" & Me.FichierXLS) = "" Then
    MsgBox "Vous devez saisir le chemin ainsi que le nom du fichier EXCEL cible", vbExclamation, "Erreur"
    Me.FichierXLS.SetFocus
    Exit Sub
End If
If UCase(Right(Me.FichierXLS, 4)) <> ".XLS" Then
    Me.FichierXLS = Me.FichierXLS & ".XLS"
End If
If Fso.FileExists(Me.FichierXLS) = True Then
    If MsgBox(Me.FichierXLS & vbCrLf & "Existe déjà voulez vous le remplacer.", vbQuestion + vbYesNo, "Fichier Existe:") = vbNo Then
        Me.FichierXLS.SetFocus
        Set Fso = Nothing
        Exit Sub
    End If
End If
Set Fso = Nothing
varProjet = Me.lstProjets
varIndice = Me.lstIndice

 Unload Me
  ExporteXls Me.FichierXLS, varProjet, varIndice
   
End Sub


Private Sub CommandButton3_Click()
Unload Me
End Sub

Private Sub Label1_Click()

End Sub

Private Sub lstProjets_Click()
Dim Rs As Recordset
Dim Sql As String
Dim indexClient As Long
Sql = "SELECT T_indiceProjet.li "
Sql = Sql & "FROM T_Projet INNER JOIN T_indiceProjet ON T_Projet.id = T_indiceProjet.IdProjet "
Sql = Sql & "WHERE T_Projet.Projet = '" & Me.lstProjets.Text & "' "
Sql = Sql & "ORDER BY T_indiceProjet.Indice;"


Set Rs = Con.OpenRecordSet(Sql)
Me.lstIndice.Clear
While Rs.EOF = False

    Me.lstIndice.AddItem Trim("" & Rs!LI)
    If Me.lstIndice.ListCount = 1 Then Me.lstIndice.Text = Trim("" & Rs!LI)
    Rs.MoveNext
Wend

End Sub

Private Sub UserForm_Layout()
Dim Rs As Recordset
Dim Sql As String
Dim indexClient As Long
Con.OpenConnetion db

Sql = "SELECT T_Projet.id, T_Projet.Projet, T_Projet.Description FROM T_Projet "
Sql = Sql & "ORDER BY  T_Projet.Projet;"

Set Rs = Con.OpenRecordSet(Sql)
While Rs.EOF = False

    Me.lstProjets.AddItem Trim("" & Rs!Projet)
    If Me.lstProjets.ListCount = 1 Then Me.lstProjets.Text = Trim("" & Rs!Projet)
    Rs.MoveNext
Wend




    


End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Con.CloseConnection
End Sub
