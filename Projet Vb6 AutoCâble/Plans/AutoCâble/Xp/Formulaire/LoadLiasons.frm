VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LoadLiasons 
   Caption         =   "Liste des liaisons manquantes :"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12930
   OleObjectBlob   =   "LoadLiasons.dsx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LoadLiasons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Noquite As Boolean
Private Sub CommandButton1_Click()
Dim sql As String
Dim Rs As Recordset
sql = "SELECT Ajout_LIAISON_CONNECTEURS.LIAISON, Ajout_LIAISON_CONNECTEURS.LIB "
sql = sql & "FROM Ajout_LIAISON_CONNECTEURS "
sql = sql & "ORDER BY Ajout_LIAISON_CONNECTEURS.LIAISON;"


Set Rs = Con.OpenRecordSet(sql)
Me.LstEcarte.Clear
Me.LstGarder.Clear
While Rs.EOF = False
    Me.LstGarder.AddItem Trim("" & Rs!Liaison)
     Me.LstGarder.List(Me.LstGarder.ListCount - 1, 1) = Trim("" & Rs!LIB)
    Rs.MoveNext
Wend


End Sub

Private Sub CommandButton2_Click()
  Dim sql As String
Dim Rs As Recordset
sql = "SELECT Ajout_LIAISON_CONNECTEURS.LIAISON, Ajout_LIAISON_CONNECTEURS.LIB "
sql = sql & "FROM Ajout_LIAISON_CONNECTEURS "
sql = sql & "ORDER BY Ajout_LIAISON_CONNECTEURS.LIAISON;"


Set Rs = Con.OpenRecordSet(sql)
Me.LstEcarte.Clear
Me.LstGarder.Clear
While Rs.EOF = False
    Me.LstEcarte.AddItem Trim("" & Rs!Liaison)
     Me.LstEcarte.List(Me.LstEcarte.ListCount - 1, 1) = Trim("" & Rs!LIB)
    Rs.MoveNext
Wend


End Sub

Private Sub CommandButton3_Click()
If Me.LstEcarte.ListIndex = -1 Then
    MsgBox "Vous devez sélectionner un élément dans la liste", vbExclamation
    Exit Sub
End If
Me.LstGarder.AddItem Me.LstEcarte.List(Me.LstEcarte.ListIndex, 0)
Me.LstGarder.List(Me.LstGarder.ListCount - 1, 1) = Me.LstEcarte.List(Me.LstEcarte.ListIndex, 1)
Me.LstEcarte.RemoveItem (Me.LstEcarte.ListIndex)
End Sub

Private Sub CommandButton4_Click()
If Me.LstGarder.ListIndex = -1 Then
    MsgBox "Vous devez sélectionner un élément dans la liste", vbExclamation
    Exit Sub
End If
Me.LstEcarte.AddItem Me.LstGarder.List(Me.LstGarder.ListIndex, 0)
Me.LstEcarte.List(Me.LstEcarte.ListCount - 1, 1) = Me.LstGarder.List(Me.LstGarder.ListIndex, 1)
Me.LstGarder.RemoveItem (Me.LstGarder.ListIndex)
End Sub

Private Sub CommandButton5_Click()
Dim sql As String
Dim Rs As Recordset
sql = "SELECT Ajout_LIAISON.LIAISON, Ajout_LIAISON.LIB "
sql = sql & "FROM Ajout_LIAISON "
sql = sql & "ORDER BY Ajout_LIAISON.LIAISON;"


Set Rs = Con.OpenRecordSet(sql)
Me.LstEcartef.Clear
Me.LstGarderF.Clear
While Rs.EOF = False
    Me.LstGarderF.AddItem Trim("" & Rs!Liaison)
     Me.LstGarderF.List(Me.LstGarderF.ListCount - 1, 1) = Trim("" & Rs!LIB)
    Rs.MoveNext
Wend


End Sub

Private Sub CommandButton6_Click()
  Dim sql As String
Dim Rs As Recordset
sql = "SELECT Ajout_LIAISON.LIAISON, Ajout_LIAISON.LIB "
sql = sql & "FROM Ajout_LIAISON "
sql = sql & "ORDER BY Ajout_LIAISON.LIAISON;"


Set Rs = Con.OpenRecordSet(sql)
Me.LstEcartef.Clear
Me.LstGarderF.Clear
While Rs.EOF = False
    Me.LstEcartef.AddItem Trim("" & Rs!Liaison)
     Me.LstEcartef.List(Me.LstEcartef.ListCount - 1, 1) = Trim("" & Rs!LIB)
    Rs.MoveNext
Wend


End Sub

Private Sub CommandButton7_Click()
If Me.LstEcartef.ListIndex = -1 Then
    MsgBox "Vous devez sélectionner un élément dans la liste", vbExclamation
    Exit Sub
End If
Me.LstGarderF.AddItem Me.LstEcartef.List(Me.LstEcartef.ListIndex, 0)
Me.LstGarderF.List(Me.LstGarderF.ListCount - 1, 1) = Me.LstEcartef.List(Me.LstEcartef.ListIndex, 1)
Me.LstEcartef.RemoveItem (Me.LstEcartef.ListIndex)
End Sub

Private Sub CommandButton8_Click()
If Me.LstGarderF.ListIndex = -1 Then
    MsgBox "Vous devez sélectionner un élément dans la liste", vbExclamation
    Exit Sub
End If
Me.LstEcartef.AddItem Me.LstGarderF.List(Me.LstGarderF.ListIndex, 0)
Me.LstEcartef.List(Me.LstEcartef.ListCount - 1, 1) = Me.LstGarderF.List(Me.LstGarderF.ListIndex, 1)
Me.LstGarderF.RemoveItem (Me.LstGarderF.ListIndex)
End Sub

Private Sub CommandButton9_Click()
Dim sql As String

If Me.LstGarder.ListCount > 0 Then
    For i = 0 To Me.LstGarder.ListCount - 1
    sql = "INSERT INTO LIAISON_CONNECTEURS ( CLIENT, LIAISON, LIB ) "
    sql = sql & "Values ( "
    sql = sql & "'" & MyReplace(Client) & "', "
    sql = sql & "'" & MyReplace(Me.LstGarder.List(i, 0)) & "', "
    sql = sql & "'" & MyReplace(Me.LstGarder.List(i, 1)) & "'"
    sql = sql & ");"
    Con.Exequte sql
    Next i
End If
If Me.LstGarderF.ListCount > 0 Then
    For i = 0 To Me.LstGarderF.ListCount - 1
    sql = "INSERT INTO LIAISON ( CLIENT, LIAISON, LIB ) "
    sql = sql & "Values ( "
    sql = sql & "'" & MyReplace(Client) & "', "
    sql = sql & "'" & MyReplace(Me.LstGarderF.List(i, 0)) & "', "
    sql = sql & "'" & MyReplace(Me.LstGarderF.List(i, 1)) & "'"
    sql = sql & ");"
    Con.Exequte sql
    Next i
End If
Noquite = False
Me.Hide
End Sub

Private Sub LstEcarte_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
CommandButton3_Click
End Sub

Private Sub LstEcartef_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
CommandButton7_Click
End Sub

Private Sub LstGarder_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
CommandButton4_Click
End Sub

Private Sub LstGarderF_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
CommandButton8_Click
End Sub

Public Sub Charger(NameClient As String)
Dim sql As String
Dim Rs As Recordset


sql = "SELECT Ajout_LIAISON_CONNECTEURS.LIAISON, Ajout_LIAISON_CONNECTEURS.LIB "
sql = sql & "FROM Ajout_LIAISON_CONNECTEURS "
sql = sql & "where  Ajout_LIAISON_CONNECTEURS.Job=" & NmJob & " "
sql = sql & "ORDER BY Ajout_LIAISON_CONNECTEURS.LIAISON;"
Set Rs = Con.OpenRecordSet(sql)
Me.LstEcarte.Clear
While Rs.EOF = False
    Me.LstEcarte.AddItem Trim("" & Rs!Liaison)
     Me.LstEcarte.List(Me.LstEcarte.ListCount - 1, 1) = Trim("" & Rs!LIB)
    Rs.MoveNext
Wend

sql = "SELECT Ajout_LIAISON.LIAISON, Ajout_LIAISON.LIB "
sql = sql & "FROM Ajout_LIAISON "
sql = sql & "where  Ajout_LIAISON.Job=" & NmJob & " "

sql = sql & "ORDER BY Ajout_LIAISON.LIAISON;"
Set Rs = Con.OpenRecordSet(sql)
Me.LstEcartef.Clear
While Rs.EOF = False
    Me.LstEcartef.AddItem Trim("" & Rs!Liaison)
     Me.LstEcartef.List(Me.LstEcartef.ListCount - 1, 1) = Trim("" & Rs!LIB)
    Rs.MoveNext
Wend

    Me.Client.Caption = NameClient
    Me.Show vbModal
    sql = "DELETE Ajout_LIAISON.*, Ajout_LIAISON.Job FROM Ajout_LIAISON "
    sql = sql & "WHERE Ajout_LIAISON.Job=" & NmJob & ";"

    Con.Exequte sql
sql = "DELETE Ajout_LIAISON_CONNECTEURS.*, Ajout_LIAISON_CONNECTEURS.Job FROM Ajout_LIAISON_CONNECTEURS "
    sql = sql & "WHERE Ajout_LIAISON_CONNECTEURS.Job=" & NmJob & ";"
    Con.Exequte sql
End Sub

Private Sub UserForm_Activate()
Noquite = True
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Cancel = Noquite
End Sub
