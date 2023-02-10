VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "Clients:"
   ClientHeight    =   8910
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

Private Sub CommandButton10_Click()
Set TableauPath = funPath
RepCom = Replace(ScanRep.chargement(RepCom), TableauPath.Item("PathServer"), "", 1)
End Sub

Private Sub CommandButton11_Click()
Set TableauPath = funPath
RepNota = Replace(ScanRep.chargement(RepNota), TableauPath.Item("PathServer"), "", 1)
End Sub

Private Sub CommandButton3_Click()
 Set TableauPath = funPath
Cartouche = Replace(ScanFichier.chargement("dwg", Cartouche), TableauPath.Item("PathServer"), "", 1)

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
Dim sql As String
Dim Rs As Recordset
Dim Msg As String
If Trim("" & Me.Client) = "" Then
    Me.Id = ""
    Exit Sub
End If

If Trim("" & Me.Id) = "" Then
    sql = "SELECT T_Clients.* FROM T_Clients WHERE T_Clients.Client='" & UCase(MyReplace(Me.Client)) & "';"
    Set Rs = Con.OpenRecordSet(sql)
    If Rs.EOF = False Then
        Msg = Me.Client & " Existe déjà"
        GoTo Err
    End If
    sql = "INSERT INTO T_Clients ( Client,Formulaire,PathConnecteurs,PathComposants,PathNotas) values('"
    sql = sql & UCase(MyReplace(Me.Client)) & "','"
    sql = sql & UCase(MyReplace(Me.Cartouche)) & "','"
    sql = sql & UCase(MyReplace(Me.RepCon)) & "','"
    sql = sql & UCase(MyReplace(Me.RepCom)) & "','"
    sql = sql & UCase(MyReplace(Me.RepNota)) & "');"
    Con.Exequte sql

Else
        sql = "SELECT T_Clients.* FROM T_Clients WHERE T_Clients.Client='" & MyReplace(Me.Client) & "' AND T_Clients.id<>" & Me.Id & ";"
    Set Rs = Con.OpenRecordSet(sql)
    If Rs.EOF = False Then
        Msg = Me.Client & " Existe déjà"
        GoTo Err
    End If
    sql = "UPDATE T_Clients SET T_Clients.Client = '" & UCase(MyReplace(Me.Client))
    sql = sql & "',Formulaire = '" & UCase(MyReplace(Me.Cartouche)) & "'"
    sql = sql & ",PathConnecteurs = '" & UCase(MyReplace(Me.RepCon)) & "'"
    sql = sql & ",PathComposants = '" & UCase(MyReplace(Me.RepCom)) & "'"
    sql = sql & ",PathNotas = '" & UCase(MyReplace(Me.RepNota)) & "'"
    sql = sql & "WHERE T_Clients.id=" & Me.Id & ";"
    Con.Exequte sql
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
Dim sql As String
Dim Rs As Recordset
Dim NbRecord As Long
If Trim("" & Me.Id) <> "" Then
    If MsgBox("Voulez vous vraiment supprimer : " & Me.Client, vbQuestion + vbYesNo) = vbNo Then Exit Sub

    sql = "SELECT T_Clients.id "
    sql = sql & "FROM T_Clients INNER JOIN T_indiceProjet ON T_Clients.Client = T_indiceProjet.Client "
    sql = sql & "WHERE T_Clients.id=" & Me.Id & ";"

    Set Rs = Con.OpenRecordSet(sql)
    If Rs.EOF = False Then
    While Rs.EOF = False
        NbRecord = NbRecord + 1
    Rs.MoveNext
    Wend
        MsgBox "Le Client : " & Me.Client & " ne peut pas être supprimé car il pointe sur " & NbRecord & " Pièce(s) Existante(s)"
        GoTo Fin
    End If
    sql = "Delete T_Clients.Client, T_Clients.Id FROM T_Clients WHERE T_Clients.id=" & Me.Id & ";"
    Con.Exequte sql
Fin:
Set Rs = Con.CloseRecordSet(Rs)
    Maj
    

End If
End Sub

Private Sub CommandButton9_Click()
Set TableauPath = funPath
RepCon = Replace(ScanRep.chargement(RepCon), TableauPath.Item("PathServer"), "", 1)
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Me.Id = Me.ListBox1.List(Me.ListBox1.ListIndex, 1)
Me.Client = Me.ListBox1.List(Me.ListBox1.ListIndex, 0)
Me.Cartouche = Me.ListBox1.List(Me.ListBox1.ListIndex, 2)
Me.RepCon = Me.ListBox1.List(Me.ListBox1.ListIndex, 3)
Me.RepCom = Me.ListBox1.List(Me.ListBox1.ListIndex, 4)
Me.RepNota = Me.ListBox1.List(Me.ListBox1.ListIndex, 5)
End Sub

Private Sub UserForm_Activate()
Maj
End Sub

Sub Maj()
Dim Rs As Recordset
Dim sql As String
Set TableauPath = funPath
 Me.ListBox1.Clear
sql = "SELECT T_Clients.Client, T_Clients.id,T_Clients.Formulaire,T_Clients.PathConnecteurs, "
sql = sql & "T_Clients.PathComposants,T_Clients.PathNotas "
sql = sql & "FROM T_Clients "
sql = sql & "ORDER BY T_Clients.Client;"

Set Rs = Con.OpenRecordSet(sql)
While Rs.EOF = False
    Me.ListBox1.AddItem Trim("" & Rs!Client)
    Me.ListBox1.List(Me.ListBox1.ListCount - 1, 1) = Trim("" & Rs!Id)
    Me.ListBox1.List(Me.ListBox1.ListCount - 1, 2) = Trim("" & Rs!Formulaire)
    Me.ListBox1.List(Me.ListBox1.ListCount - 1, 3) = Trim("" & Rs!PathConnecteurs)
    Me.ListBox1.List(Me.ListBox1.ListCount - 1, 4) = Trim("" & Rs!PathComposants)
    Me.ListBox1.List(Me.ListBox1.ListCount - 1, 5) = Trim("" & Rs!PathNotas)
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
Me.Id = ""
End Sub
