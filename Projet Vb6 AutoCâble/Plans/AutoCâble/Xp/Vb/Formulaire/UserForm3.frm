VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "Clients:"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   Icon            =   "UserForm3.dsx":0000
   OleObjectBlob   =   "UserForm3.dsx":08CA
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton10_Click()
Set TableauPath = funPath
RepCom = Replace(ScanRep.Chargement(RepCom), TableauPath.Item("PathServer"), "", 1)
End Sub

Private Sub CommandButton11_Click()
Set TableauPath = funPath
RepNota = Replace(ScanRep.Chargement(RepNota), TableauPath.Item("PathServer"), "", 1)
End Sub

Private Sub CommandButton12_Click()
Set TableauPath = funPath
Catalogue = Replace(ScanFichier.Chargement("mdb", Catalogue), TableauPath.Item("PathServer"), "", 1)

End Sub

Private Sub CommandButton3_Click()
 Set TableauPath = funPath
Cartouche = Replace(ScanFichier.Chargement("dwg", Cartouche), TableauPath.Item("PathServer"), "", 1)

End Sub

Private Sub CommandButton8_Click()
Unload Me
End Sub

Private Sub ListBox2_Click()

End Sub

Private Sub ListBox2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
 'Me.LibEnsemble = Me.LstEnsemble.List(Me.LstEnsemble.ListIndex, 0)
 'Me.IdEnsemble = Me.LstEnsemble.List(Me.LstEnsemble.ListIndex, 1)

End Sub

Private Sub CommandButton5_Click()
Dim Sql As String
Dim Rs As Recordset
Dim Msg As String
If Trim("" & Me.Client) = "" Then
    Me.Id = ""
    Exit Sub
End If

If Trim("" & Me.Id) = "" Then
    Sql = "SELECT T_Clients.* FROM T_Clients WHERE T_Clients.Client='" & UCase(MyReplace(Me.Client)) & "';"
    Set Rs = Con.OpenRecordSet(Sql)
    If Rs.EOF = False Then
        Msg = Me.Client & " Existe déjà"
        GoTo Err
    End If
    Sql = "INSERT INTO T_Clients ( Client,Formulaire,PathConnecteurs,PathComposants,PathNotas,PathCatalogue) values('"
    Sql = Sql & UCase(MyReplace(Me.Client)) & "','"
    Sql = Sql & UCase(MyReplace(Me.Cartouche)) & "','"
    Sql = Sql & UCase(MyReplace(Me.RepCon)) & "','"
    Sql = Sql & UCase(MyReplace(Me.RepCom)) & "','"
    Sql = Sql & UCase(MyReplace(Me.RepNota)) & "','"
    Sql = Sql & UCase(MyReplace(Me.Catalogue)) & "')"
    Con.Exequte Sql

Else
        Sql = "SELECT T_Clients.* FROM T_Clients WHERE T_Clients.Client='" & MyReplace(Me.Client) & "' AND T_Clients.id<>" & Me.Id & ";"
    Set Rs = Con.OpenRecordSet(Sql)
    If Rs.EOF = False Then
        Msg = Me.Client & " Existe déjà"
        GoTo Err
    End If
    Sql = "UPDATE T_Clients SET T_Clients.Client = '" & UCase(MyReplace(Me.Client))
    Sql = Sql & "',Formulaire = '" & UCase(MyReplace(Me.Cartouche)) & "'"
    Sql = Sql & ",PathConnecteurs = '" & UCase(MyReplace(Me.RepCon)) & "'"
    Sql = Sql & ",PathComposants = '" & UCase(MyReplace(Me.RepCom)) & "'"
    Sql = Sql & ",PathNotas = '" & UCase(MyReplace(Me.RepNota)) & "'"
    Sql = Sql & ",PathCatalogue = '" & UCase(MyReplace(Me.Catalogue)) & "'"
    
    Sql = Sql & "WHERE T_Clients.id=" & Me.Id & ";"
    Con.Exequte Sql
End If



GoTo Fin
Err:
MsgBox Msg
Fin:
Set Rs = Con.CloseRecordSet(Rs)
Maj

End Sub

Private Sub CommandButton6_Click()
Maj
End Sub

Private Sub CommandButton7_Click()
Dim Sql As String
Dim Rs As Recordset
Dim NbRecord As Long
If Trim("" & Me.Id) <> "" Then
    If MsgBox("Voulez vous vraiment supprimer : " & Me.Client, vbQuestion + vbYesNo) = vbNo Then Exit Sub

    Sql = "SELECT T_Clients.id "
    Sql = Sql & "FROM T_Clients INNER JOIN T_indiceProjet ON T_Clients.Client = T_indiceProjet.Client "
    Sql = Sql & "WHERE T_Clients.id=" & Me.Id & ";"

    Set Rs = Con.OpenRecordSet(Sql)
    If Rs.EOF = False Then
    While Rs.EOF = False
        NbRecord = NbRecord + 1
    Rs.MoveNext
    Wend
        MsgBox "Le Client : " & Me.Client & " ne peut pas être supprimé car il pointe sur " & NbRecord & " Pièce(s) Existante(s)"
        GoTo Fin
    End If
    Sql = "Delete T_Clients.Client, T_Clients.Id FROM T_Clients WHERE T_Clients.id=" & Me.Id & ";"
    Con.Exequte Sql
Fin:
Set Rs = Con.CloseRecordSet(Rs)
    Maj
    

End If
End Sub

Private Sub CommandButton9_Click()
Set TableauPath = funPath
RepCon = Replace(ScanRep.Chargement(RepCon), TableauPath.Item("PathServer"), "", 1)
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Me.Id = Me.ListBox1.List(Me.ListBox1.ListIndex, 1)
Me.Client = Me.ListBox1.List(Me.ListBox1.ListIndex, 0)
Me.Cartouche = Me.ListBox1.List(Me.ListBox1.ListIndex, 2)
Me.RepCon = Me.ListBox1.List(Me.ListBox1.ListIndex, 3)
Me.RepCom = Me.ListBox1.List(Me.ListBox1.ListIndex, 4)
Me.RepNota = Me.ListBox1.List(Me.ListBox1.ListIndex, 5)
Me.Catalogue = Me.ListBox1.List(Me.ListBox1.ListIndex, 6)
End Sub

Private Sub UserForm_Activate()
Maj
End Sub

Sub Maj()
Dim Rs As Recordset
Dim Sql As String
Set TableauPath = funPath
 Me.ListBox1.Clear
Sql = "SELECT T_Clients.Client, T_Clients.id,T_Clients.Formulaire,T_Clients.PathConnecteurs, "
Sql = Sql & "T_Clients.PathComposants,T_Clients.PathNotas,T_Clients.PathCatalogue "
Sql = Sql & "FROM T_Clients "
Sql = Sql & "ORDER BY T_Clients.Client;"

Set Rs = Con.OpenRecordSet(Sql)
While Rs.EOF = False
    Me.ListBox1.AddItem Trim("" & Rs!Client)
    Me.ListBox1.List(Me.ListBox1.ListCount - 1, 1) = Trim("" & Rs!Id)
    Me.ListBox1.List(Me.ListBox1.ListCount - 1, 2) = Trim("" & Rs!Formulaire)
    Me.ListBox1.List(Me.ListBox1.ListCount - 1, 3) = Trim("" & Rs!PathConnecteurs)
    Me.ListBox1.List(Me.ListBox1.ListCount - 1, 4) = Trim("" & Rs!PathComposants)
    Me.ListBox1.List(Me.ListBox1.ListCount - 1, 5) = Trim("" & Rs!PathNotas)
    Me.ListBox1.List(Me.ListBox1.ListCount - 1, 6) = Trim("" & Rs!PathCatalogue)
        If Me.ListBox1.ListCount = 1 Then Me.ListBox1.ListIndex = 0

    Rs.MoveNext
Wend
Set Rs = Con.CloseRecordSet(Rs)
Me.Frame4.Enabled = Admin
Me.Client = ""
Me.RepCom = ""
Me.RepCon = ""
Me.RepNota = ""
Me.Cartouche = ""
Me.Catalogue = ""
Me.Id = ""
End Sub
