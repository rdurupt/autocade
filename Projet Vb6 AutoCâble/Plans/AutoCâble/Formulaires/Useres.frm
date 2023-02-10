VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Useres 
   Caption         =   "Login:"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3465
   OleObjectBlob   =   "Useres.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Useres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CmdAnnuler_Click()
 NoClose = False
boolExec = False
 boolQuitte = True
 Unload Me

End Sub

Private Sub CmdLogin_Click()
Dim Sql As String
Dim RsUser As Recordset
If Trim("" & Me.UsersName) = "" Then
    MsgBox "Le champ User Name est obligatoire", vbExclamation, "Login"
     Me.UsersName.SetFocus
     Exit Sub
End If
If Trim("" & Me.PasseWord) = "" Then
    MsgBox "Le champ Pass word est obligatoire", vbExclamation, "Login"
     Me.PasseWord.SetFocus
     Exit Sub
End If
Sql = "SELECT T_Users.User, T_Users.PassWord, T_Users.Cloturer, T_Groupe.Admin, T_Groupe.Lecture, T_Groupe.Ecriture, T_Groupe.Creation "
Sql = Sql & "FROM T_Groupe INNER JOIN T_Users ON T_Groupe.id = T_Users.IdGoupe "
Sql = Sql & "WHERE T_Users.User='" & Me.UsersName & "' "
Sql = Sql & "AND T_Users.PassWord ='" & Me.PasseWord & "';"
Con.OpenConnetion db

Set RsUser = Con.OpenRecordSet(Sql)

If RsUser.EOF = False Then
    If RsUser!Cloturer = True Then
        MsgBox "Votre compte est verrouillé.", vbExclamation, "Login"
         Con.CloseConnection
        Exit Sub
    End If
    Loguer = True
    Admin = RsUser!Admin
    Lecture = RsUser!Lecture
    Ecriture = RsUser!Ecriture
    Creation = RsUser!Creation
   
    Con.CloseConnection
     NoClose = False

  Unload Me
'  Load Menu
    Menu.Show
  
    Exit Sub
Else
    MsgBox "User Name ou Pass word inexistant vérifier ces paramètres ou contactez l'administrateur de l'application !", vbCritical, "Login"
    
End If
    Con.CloseConnection

End Sub

Private Sub UserForm_Activate()
NoClose = True

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Cancel = NoClose

End Sub
