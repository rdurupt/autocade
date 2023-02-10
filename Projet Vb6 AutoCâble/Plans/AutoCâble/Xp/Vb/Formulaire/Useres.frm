VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Useres 
   Caption         =   "Login:"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3465
   Icon            =   "Useres.dsx":0000
   OleObjectBlob   =   "Useres.dsx":030A
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Useres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BoolErr As Boolean
Dim txtForm As String
Dim MyDroits As String
Dim Noquite As Boolean
Public Sub Charger(Mytype As String, Droits As String, Optional ModuleOk As Boolean)
DoEvents
txtForm = Mytype
MyDroits = Droits
DoEvents
Dim Sql As String
Dim Rs As Recordset
Sql = "SELECT T_Users.User FROM T_Users ORDER BY T_Users.User;"

Set Rs = Con.OpenRecordSet(Sql)
Me.UsersName.Clear
While Rs.EOF = False
Me.UsersName.AddItem "" & Rs!User
DoEvents
    Rs.MoveNext
Wend
If Me.UsersName.ListCount > 0 Then Me.UsersName.ListIndex = 0
NoClose = True
Set Rs = Con.CloseRecordSet(Rs)
DoEvents
Me.Show vbModal

End Sub

Private Sub CmdAnnuler_Click()
 NoClose = False
boolExec = False
 boolQuitte = True
 Me.Hide

End Sub

Private Sub CmdLogin_Click()
Dim Sql As String
Dim RsUser As Recordset
Dim FiltreDroits As String
BoolErr = False
If Trim("" & Me.UsersName) = "" Then
    MsgBox "Le champ User Name est obligatoire", vbExclamation, "Login"
    BoolErr = True
     Me.UsersName.SetFocus
     Exit Sub
End If
If Trim("" & Me.PasseWord) = "" Then
    MsgBox "Le champ Pass word est obligatoire", vbExclamation, "Login"
    BoolErr = True
     Me.PasseWord.SetFocus
     Exit Sub
End If
    
Select Case MyDroits
    Case "Admin"
        FiltreDroits = "Admin=True "
    Case "Vérificateur"
         FiltreDroits = "Verificateur=True "
    Case "Approbateur"
         FiltreDroits = "Approbateur=True"
End Select
Sql = "SELECT T_Users.User, T_Users.PassWord, T_Users.Cloturer,T_Users.Admin, "
Sql = Sql & "T_Users.Verificateur, T_Users.Approbateur "
Sql = Sql & "FROM T_Users "
Sql = Sql & "WHERE T_Users.User='" & Me.UsersName & "' "
Sql = Sql & "AND T_Users.PassWord ='" & Me.PasseWord & "' ;"


Set RsUser = Con.OpenRecordSet(Sql)

If RsUser.EOF = False Then
    If RsUser!Cloturer = True Then
        MsgBox "Votre compte est verrouillé.", vbExclamation, "Login"
         
        Exit Sub
    End If
RsUser.Filter = FiltreDroits
If RsUser.EOF = True Then
    MsgBox "Vous n 'avez pas les droit (" & MyDroits & ")", vbExclamation, "Login"
    
    Exit Sub
End If

    Loguer = True
    Admin = RsUser!Admin
     Verifrificateur = RsUser!Verificateur
 Approbateur = RsUser!Approbateur
'    Lecture = RsUser!Lecture
'    Ecriture = RsUser!Ecriture
'    Creation = RsUser!Creation
     Set RsUser = Con.CloseRecordSet(RsUser)
    
     NoClose = False

  
'  Load Menu
Select Case txtForm
    Case "NULL"
         Me.Hide
    Case "Creer"
            Me.Hide
           Creer.Chargement
           Unload Creer
    Case "MenuAdmin"
        Me.Hide
        MenuAdmin.Show vbModal
        Unload MenuAdmin
     Case "Vérificateur"
          Me.Hide
          VerifierEtude.Show vbModal
          Unload VerifierEtude
    Case "Approbation"
          Me.Hide
          Approbation.Show vbModal
          Unload Approbation
             
    Case "ModifierCartouches"
        Me.Hide
        ModifierCartouches.Show vbModal
        Unload ModifierCartouches
        
    Case "ImportCablePrix"
        Me.Hide
        ImportCablePrixExport.ImporOk = True
        ImportCablePrixExport.Show vbModal
        Unload ImportCablePrixExport
        
     Case "ExportCablePrix"
        Me.Hide
        ImportCablePrixExport.ImporOk = False
        ImportCablePrixExport.Show vbModal
        Unload ImportCablePrixExport
End Select

   
NoClose = False

'  ImportCablePrix  ImportCablePrixExport
Else
    MsgBox "Pass word erroné. Vérifiez ce paramètre ou contactez l'administrateur de l'application.", vbCritical, "Login"
    
End If
'Set RsUser = Con.CloseRecordSet(RsUser)
    
  

End Sub

Private Sub CmdLogin_Enter()
CmdLogin_Click
 If BoolErr = True Then
     Me.PasseWord.SetFocus
    DoEvents
 End If
End Sub





Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Cancel = NoClose

End Sub
