VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Useres 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Login:"
   ClientHeight    =   1785
   ClientLeft      =   30
   ClientTop       =   195
   ClientWidth     =   3465
   Icon            =   "Useres.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OleObjectBlob   =   "Useres.dsx":030A
   ShowInTaskbar   =   0   'False
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
Dim MyModuleOk As Boolean
Public DroitsOk As Boolean
'Public Sub charger(Mytype As String, Optional Droits As String, Optional ModuleOk As Boolean)
'DoEvents
'txtForm = Mytype
'MyDroits = Droits
'DoEvents
'Dim sql As String
'Dim Rs As Recordset
'sql = "SELECT T_Users.User FROM T_Users ORDER BY T_Users.User;"
'
'Set Rs = Con.OpenRecordSet(sql)
'Me.UsersName.Clear
'While Rs.EOF = False
'Me.UsersName.AddItem "" & Rs!User
'DoEvents
'    Rs.MoveNext
'Wend
'If Me.UsersName.ListCount > 0 Then Me.UsersName.ListIndex = 0
'NoClose = True
'Set Rs = Con.CloseRecordSet(Rs)
'DoEvents
'Me.Show vbModal
'
'End Sub
Public Sub charger(Mytype As String, Optional Droits As String, Optional ModuleOk As Boolean)
Dim UserDefault As String
Dim UserPassDefault As String
DoEvents
txtForm = Mytype
MyDroits = Droits
DoEvents
MyModuleOk = ModuleOk
Dim sql As String
Dim Rs As Recordset
 UserDefault = GetDefault("UserDefault", "")
 UserPassDefault = GetDefault("UserPassDefault", "")
sql = "SELECT T_Users.User FROM T_Users ORDER BY T_Users.User;"

Set Rs = Con.OpenRecordSet(sql)
Me.UsersName.Clear
While Rs.EOF = False
Debug.Print "" & Rs!User
Me.UsersName.AddItem "" & Rs!User
If UserDefault = "" & Rs!User Then
    Me.UsersName.ListIndex = Me.UsersName.ListCount - 1
    If Trim("" & UserPassDefault) <> "" Then Me.PasseWord = UserPassDefault
End If
DoEvents
    Rs.MoveNext
Wend
If Me.UsersName.ListCount > 0 And Me.UsersName.ListIndex = -1 Then Me.UsersName.ListIndex = 0
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
Dim sql As String
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
    sql = "SELECT  T_Users.Id as Id_Users,T_Droits.Id_Bouton, T_Users.User, T_Users.PassWord, T_Users.Cloturer "
sql = sql & "FROM (T_Boutons INNER JOIN T_Droits ON T_Boutons.Id = T_Droits.Id_Bouton) LEFT JOIN T_Users ON T_Droits.Id_Useur = T_Users.Id "
sql = sql & "WHERE T_Boutons.Name='" & MyReplace(txtForm) & "' "
sql = sql & "and T_Users.Cloturer= false "
sql = sql & "and User='" & MyReplace(UsersName) & "' "
sql = sql & "and T_Users.PassWord='" & MyReplace(PasseWord) & "';"
Set RsUser = Con.OpenRecordSet(sql)
If RsUser.EOF = True Then
    MsgBox "Vous n'avez pas les droits sur : " & txtForm & vbCrLf & "ou Votre  Pass Word est erroné" & vbCrLf & "et/ou Votre compte a été verrouillé."
    Set RsUser = Con.CloseRecordSet(RsUser)
    Exit Sub
 Else
    If Trim(PasseWord) <> RsUser!PassWord Then
           MsgBox "Vous n'avez pas les droits sur : " & txtForm & vbCrLf & "ou Votre  Pass Word est erroné" & vbCrLf & "et/ou Votre compte a été verrouillé."
    Set RsUser = Con.CloseRecordSet(RsUser)
    Exit Sub
    Else
        If RsUser!Cloturer = True Then
            MsgBox "Votre compte est verrouillé.", vbExclamation, "Login"
            Set RsUser = Con.CloseRecordSet(RsUser)
            Exit Sub
        End If
    End If
 DroitsOk = True
 If MyModuleOk = False Then
 Id_Users = RsUser!Id_Users
 End If
  Set RsUser = Con.CloseRecordSet(RsUser)
'  frmAutocâble.DesEnabledMenu
  MajDroitsFrm Id_Users
End If

NoClose = False
Me.Hide
'Select Case MyDroits
'    Case "Admin"
'        FiltreDroits = "Admin=True "
'    Case "Vérificateur"
'         FiltreDroits = "Verificateur=True "
'    Case "Approbateur"
'         FiltreDroits = "Approbateur=True"
'End Select
'Sql = "SELECT T_Users.User, T_Users.PassWord, T_Users.Cloturer,T_Users.Admin, "
'Sql = Sql & "T_Users.Verificateur, T_Users.Approbateur "
'Sql = Sql & "FROM T_Users "
'Sql = Sql & "WHERE T_Users.User='" & Me.UsersName & "' "
'Sql = Sql & "AND T_Users.PassWord ='" & Me.PasseWord & "' ;"
'
'
'Set RsUser = Con.OpenRecordSet(Sql)
'
'If RsUser.EOF = False Then
'    If RsUser!Cloturer = True Then
'        MsgBox "Votre compte est verrouillé.", vbExclamation, "Login"
'
'        Exit Sub
'    End If
'RsUser.Filter = FiltreDroits
'If RsUser.EOF = True Then
'    MsgBox "Vous n 'avez pas les droit (" & MyDroits & ")", vbExclamation, "Login"
'
'    Exit Sub
'End If
'
'    Loguer = True
'    Admin = RsUser!Admin
'     Verifrificateur = RsUser!Verificateur
' Approbateur = RsUser!Approbateur
''    Lecture = RsUser!Lecture
''    Ecriture = RsUser!Ecriture
''    Creation = RsUser!Creation
'     Set RsUser = Con.CloseRecordSet(RsUser)
'
'     NoClose = False
'
'
''  Load Menu
'Select Case txtForm
'    Case "NULL"
'         Me.Hide
'    Case "Creer"
'            Me.Hide
'           Creer.Chargement
'           Unload Creer
'    Case "MenuAdmin"
'        Me.Hide
'        MenuAdmin.Show vbmodal
'        Unload MenuAdmin
'     Case "Vérificateur"
'          Me.Hide
'          VerifierEtude.Show vbmodal
'          Unload VerifierEtude
'    Case "Approbation"
'          Me.Hide
'          Approbation.Show vbmodal
'          Unload Approbation
'
'    Case "ModifierCartouches"
'        Me.Hide
'        ModifierCartouches.Show vbmodal
'        Unload ModifierCartouches
'
'    Case "ImportCablePrix"
'        Me.Hide
'        ImportCablePrixExport.ImporOk = True
'        ImportCablePrixExport.Show vbmodal
'        Unload ImportCablePrixExport
'
'     Case "ExportCablePrix"
'        Me.Hide
'        ImportCablePrixExport.ImporOk = False
'        ImportCablePrixExport.Show vbmodal
'        Unload ImportCablePrixExport
'End Select
'
'
'NoClose = False
'
''  ImportCablePrix  ImportCablePrixExport
'Else
'    MsgBox "Pass word erroné. Vérifiez ce paramètre ou contactez l'administrateur de l'application.", vbCritical, "Login"
'
'End If
''Set RsUser = Con.CloseRecordSet(RsUser)
'
  

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
