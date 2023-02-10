VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "Clients:"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton8_Click()
Unload Me
End Sub

Private Sub ListBox2_Click()

End Sub

Private Sub ListBox2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
 Me.LibEnsemble = Me.LstEnsemble.List(Me.LstEnsemble.ListIndex, 0)
 Me.IdEnsemble = Me.LstEnsemble.List(Me.LstEnsemble.ListIndex, 1)

End Sub

Private Sub CommandButton5_Click()
Dim Sql As String
Dim Rs As Recordset
Dim MSG As String
If Trim("" & Me.Client) = "" Then
    Me.Id = ""
    Exit Sub
End If
Con.OpenConnetion db
If Trim("" & Me.Id) = "" Then
    Sql = "SELECT T_Clients.* FROM T_Clients WHERE T_Clients.Client='" & UCase(MyReplace(Me.Client)) & "';"
    Set Rs = Con.OpenRecordSet(Sql)
    If Rs.EOF = False Then
        MSG = Me.Client & " Existe déjà"
        GoTo Err
    End If
    Sql = "INSERT INTO T_Clients ( Client ) values('" & UCase(MyReplace(Me.Client)) & "');"
    Con.Exequte Sql

Else
        Sql = "SELECT T_Clients.* FROM T_Clients WHERE T_Clients.Client='" & MyReplace(Me.Client) & "' AND T_Clients.id<>" & Me.Id & ";"
    Set Rs = Con.OpenRecordSet(Sql)
    If Rs.EOF = False Then
        MSG = Me.Client & " Existe déjà"
        GoTo Err
    End If
    Sql = "UPDATE T_Clients SET T_Clients.Client = '" & UCase(MyReplace(Me.Client)) & "' WHERE T_Clients.id=" & Me.Id & ";"
    Con.Exequte Sql
End If
Me.Id = ""
Maj
Set Rs = Con.CloseRecordSet(Rs)
Con.OpenConnetion db
Exit Sub
Err:
MsgBox MSG
Me.Id = ""
Me.Client = ""
Set Rs = Con.CloseRecordSet(Rs)
Con.OpenConnetion db

End Sub

Private Sub CommandButton6_Click()
Me.Id = ""
Me.Client = ""
End Sub

Private Sub CommandButton7_Click()
Dim Sql As String
Dim Rs As Recordset
Con.OpenConnetion db
If Trim("" & Me.Id) Then
    If MsgBox("Voulez vous vraiment supprimer : " & Me.Client, vbQuestion + vbYesNo) = vbNo Then Exit Sub

    Sql = "SELECT T_Clients.Client, T_Clients.id, T_Clients.Formulaire "
    Sql = Sql & "FROM T_Clients "
    Sql = Sql & "WHERE T_Clients.id=" & Me.Id & " "
    Sql = Sql & "AND T_Clients.Formulaire Is Not Null;"
    Set Rs = Con.OpenRecordSet(Sql)
    If Rs.EOF = False Then
        MsgBox "Le Client : " & Me.Client & " ne peut pas être supprimé car il pointe sur un Objet Système"
        GoTo Fin
    End If
    Sql = "Delete T_Clients.Client, T_Clients.Id FROM T_Clients WHERE T_Clients.id=" & Me.Id & ";"
    Con.Exequte Sql
Fin:
Set Rs = Con.CloseRecordSet(Rs)
    Maj
    Con.CloseConnection
Me.Id = ""
Me.Client = ""
End If
End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Me.Id = Me.ListBox1.List(Me.ListBox1.ListIndex, 1)
Me.Client = Me.ListBox1.List(Me.ListBox1.ListIndex, 0)
End Sub

Private Sub UserForm_Activate()
Dim Rs As Recordset
Dim Sql As String
Con.OpenConnetion db
Sql = "SELECT T_Clients.Client, T_Clients.id "
Sql = Sql & "FROM T_Clients "
Sql = Sql & "ORDER BY T_Clients.Client;"
Con.OpenConnetion db
Set Rs = Con.OpenRecordSet(Sql)
While Rs.EOF = False
    Me.ListBox1.AddItem Trim("" & Rs!Client)
    Me.ListBox1.List(Me.ListBox1.ListCount - 1, 1) = Trim("" & Rs!Id)
        If Me.ListBox1.ListCount = 1 Then Me.ListBox1.ListIndex = 0

    Rs.MoveNext
Wend
Set Rs = Con.CloseRecordSet(Rs)
Con.CloseConnection
End Sub

Private Sub UserForm_Click()

End Sub
Sub Maj()
Dim Rs As Recordset
Dim Sql As String
Me.Client = ""
 Me.ListBox1.Clear
Sql = "SELECT T_Clients.Client, T_Clients.id "
Sql = Sql & "FROM T_Clients "
Sql = Sql & "ORDER BY T_Clients.Client;"

Set Rs = Con.OpenRecordSet(Sql)
While Rs.EOF = False
    Me.ListBox1.AddItem Trim("" & Rs!Client)
    Me.ListBox1.List(Me.ListBox1.ListCount - 1, 1) = Trim("" & Rs!Id)
        If Me.ListBox1.ListCount = 1 Then Me.ListBox1.ListIndex = 0

    Rs.MoveNext
Wend
Set Rs = Con.CloseRecordSet(Rs)
End Sub
