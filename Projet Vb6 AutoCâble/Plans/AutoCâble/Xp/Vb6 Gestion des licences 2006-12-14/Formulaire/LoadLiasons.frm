VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LoadLiasons 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Liste des liaisons manquantes :"
   ClientHeight    =   8985
   ClientLeft      =   30
   ClientTop       =   300
   ClientWidth     =   12930
   Icon            =   "LoadLiasons.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OleObjectBlob   =   "LoadLiasons.dsx":08CA
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "LoadLiasons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Noquite As Boolean
Private Sub CommandButton1_Click()
Dim Sql As String
Dim Rs As Recordset
Sql = "SELECT Ajout_LIAISON_CONNECTEURS.LIAISON, Ajout_LIAISON_CONNECTEURS.LIB "
Sql = Sql & "FROM Ajout_LIAISON_CONNECTEURS "
Sql = Sql & "ORDER BY Ajout_LIAISON_CONNECTEURS.LIAISON;"


Set Rs = Con.OpenRecordSet(Sql)
Me.LstEcarte.Clear
Me.lstGarder.Clear
While Rs.EOF = False
    Me.lstGarder.AddItem Trim("" & Rs!Liaison)
     Me.lstGarder.List(Me.lstGarder.ListCount - 1, 1) = Trim("" & Rs!Lib)
    Rs.MoveNext
Wend


End Sub

Private Sub CommandButton2_Click()
  Dim Sql As String
Dim Rs As Recordset
Sql = "SELECT Ajout_LIAISON_CONNECTEURS.LIAISON, Ajout_LIAISON_CONNECTEURS.LIB "
Sql = Sql & "FROM Ajout_LIAISON_CONNECTEURS "
Sql = Sql & "ORDER BY Ajout_LIAISON_CONNECTEURS.LIAISON;"


Set Rs = Con.OpenRecordSet(Sql)
Me.LstEcarte.Clear
Me.lstGarder.Clear
While Rs.EOF = False
    Me.LstEcarte.AddItem Trim("" & Rs!Liaison)
     Me.LstEcarte.List(Me.LstEcarte.ListCount - 1, 1) = Trim("" & Rs!Lib)
    Rs.MoveNext
Wend


End Sub

Private Sub CommandButton3_Click()
If Me.LstEcarte.ListIndex = -1 Then
    MsgBox "Vous devez sélectionner un élément dans la liste", vbExclamation
    Exit Sub
End If
Me.lstGarder.AddItem Me.LstEcarte.List(Me.LstEcarte.ListIndex, 0)
Me.lstGarder.List(Me.lstGarder.ListCount - 1, 1) = Me.LstEcarte.List(Me.LstEcarte.ListIndex, 1)
Me.LstEcarte.RemoveItem (Me.LstEcarte.ListIndex)
End Sub

Private Sub CommandButton4_Click()
If Me.lstGarder.ListIndex = -1 Then
    MsgBox "Vous devez sélectionner un élément dans la liste", vbExclamation
    Exit Sub
End If
Me.LstEcarte.AddItem Me.lstGarder.List(Me.lstGarder.ListIndex, 0)
Me.LstEcarte.List(Me.LstEcarte.ListCount - 1, 1) = Me.lstGarder.List(Me.lstGarder.ListIndex, 1)
Me.lstGarder.RemoveItem (Me.lstGarder.ListIndex)
End Sub

Private Sub CommandButton5_Click()
Dim Sql As String
Dim Rs As Recordset
Sql = "SELECT Ajout_LIAISON.LIAISON, Ajout_LIAISON.LIB "
Sql = Sql & "FROM Ajout_LIAISON "
Sql = Sql & "ORDER BY Ajout_LIAISON.LIAISON;"


Set Rs = Con.OpenRecordSet(Sql)
Me.LstEcartef.Clear
Me.LstGarderF.Clear
While Rs.EOF = False
    Me.LstGarderF.AddItem Trim("" & Rs!Liaison)
     Me.LstGarderF.List(Me.LstGarderF.ListCount - 1, 1) = Trim("" & Rs!Lib)
    Rs.MoveNext
Wend


End Sub

Private Sub CommandButton6_Click()
  Dim Sql As String
Dim Rs As Recordset
Sql = "SELECT Ajout_LIAISON.LIAISON, Ajout_LIAISON.LIB "
Sql = Sql & "FROM Ajout_LIAISON "
Sql = Sql & "ORDER BY Ajout_LIAISON.LIAISON;"


Set Rs = Con.OpenRecordSet(Sql)
Me.LstEcartef.Clear
Me.LstGarderF.Clear
While Rs.EOF = False
    Me.LstEcartef.AddItem Trim("" & Rs!Liaison)
     Me.LstEcartef.List(Me.LstEcartef.ListCount - 1, 1) = Trim("" & Rs!Lib)
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
Dim Sql As String
Dim MyExcel As New EXCEL.Application
Dim MyWorkbook As Workbook
Dim MySheet As Worksheet
Dim MalOk As Boolean
'MyExcel.Visible = True
Set MyWorkbook = MyExcel.Workbooks.Add
 Set MySheet = IsertSheet(MyWorkbook, "NULL")
If Me.lstGarder.ListCount > 0 Then
    Set MySheet = IsertSheet(MyWorkbook, "LIAISON_CONNECTEURS")
    
    For I = 0 To Me.lstGarder.ListCount - 1
    MySheet.Cells(1, 1) = "CLIENT"
    MySheet.Cells(1, 2) = "LIAISON"
    MySheet.Cells(1, 3) = "LIB"
    
    MySheet.Cells(I + 2, 1) = Client
     MySheet.Cells(I + 2, 2) = Me.lstGarder.List(I, 0)
    MySheet.Cells(I + 2, 3) = Me.lstGarder.List(I, 1)

    Next I
End If
If Me.LstGarderF.ListCount > 0 Then
Set MySheet = IsertSheet(MyWorkbook, "LIAISON_FILS")
    For I = 0 To Me.LstGarderF.ListCount - 1
    
    MySheet.Cells(1, 1) = "CLIENT"
    MySheet.Cells(1, 2) = "LIAISON"
    MySheet.Cells(1, 3) = "LIB"
    Sql = Sql & "Values ( "
    MySheet.Cells(I + 2, 1) = Client
     MySheet.Cells(I + 2, 2) = Me.LstGarderF.List(I, 0)
    MySheet.Cells(I + 2, 3) = Me.LstGarderF.List(I, 1)
    
    
   
    Next I
End If
If MySheet.Name <> "NULL" Then
 Set MySheet = IsertSheet(MyWorkbook, "NULL")
 MySheet.Delete
Dim Fso As New FileSystemObject

MalOk = True
Dim Fil As String
Fil = Environ("USERPROFILE") & "\Mes Documents\Liason_(Machine_" & Machine & ")_" & Format(Now, "yyyy-mm-dd_hh-mm-ss")
If Fso.FileExists(Fil & ".XLS") = True Then
    Fso.DeleteFile Fil & ".XLS"
End If
    
    MyWorkbook.SaveAs Fil
    
End If
MyWorkbook.Close False
MyExcel.Quit
Set MyWorkbook = Nothing
Set MyExcel = Nothing
If MalOk = True Then
    SendMal "Ajouter Liaisons", Fil & ".XLS"
    Fso.DeleteFile Fil & ".XLS"
End If
Set Fso = Nothing
Noquite = False
Me.Hide
End Sub

Private Sub LstEcarte_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
CommandButton3_Click
End Sub

Private Sub LstEcartef_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
CommandButton7_Click
End Sub

Private Sub lstGarder_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
CommandButton4_Click
End Sub

Private Sub LstGarderF_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
CommandButton8_Click
End Sub

Public Sub charger(NameClient As String)
Dim Sql As String
Dim Rs As Recordset
Sql = "SELECT T_Serveur_Smtp.Activer FROM T_Serveur_Smtp WHERE T_Serveur_Smtp.Activer=True;"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = True Then GoTo Fin

Sql = "SELECT Ajout_LIAISON_CONNECTEURS.LIAISON, Ajout_LIAISON_CONNECTEURS.LIB "
Sql = Sql & "FROM Ajout_LIAISON_CONNECTEURS "
Sql = Sql & "where  Ajout_LIAISON_CONNECTEURS.Job=" & NmJob & " "
Sql = Sql & "ORDER BY Ajout_LIAISON_CONNECTEURS.LIAISON;"
Set Rs = Con.OpenRecordSet(Sql)
Me.LstEcarte.Clear
While Rs.EOF = False
    Me.LstEcarte.AddItem Trim("" & Rs!Liaison)
     Me.LstEcarte.List(Me.LstEcarte.ListCount - 1, 1) = Trim("" & Rs!Lib)
    Rs.MoveNext
Wend

Sql = "SELECT Ajout_LIAISON.LIAISON, Ajout_LIAISON.LIB "
Sql = Sql & "FROM Ajout_LIAISON "
Sql = Sql & "where  Ajout_LIAISON.Job=" & NmJob & " "

Sql = Sql & "ORDER BY Ajout_LIAISON.LIAISON;"
Set Rs = Con.OpenRecordSet(Sql)
Me.LstEcartef.Clear
While Rs.EOF = False
    Me.LstEcartef.AddItem Trim("" & Rs!Liaison)
     Me.LstEcartef.List(Me.LstEcartef.ListCount - 1, 1) = Trim("" & Rs!Lib)
    Rs.MoveNext
Wend

    Me.Client.Caption = NameClient
    
    Me.Show vbModal
   
Fin:
   Set Rs = Con.CloseRecordSet(Rs)
End Sub

Private Sub UserForm_Activate()
Noquite = True
Me.Repaint
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Dim Sql As String
Cancel = Noquite
 Sql = "DELETE Ajout_LIAISON.*, Ajout_LIAISON.Job FROM Ajout_LIAISON "
    Sql = Sql & "WHERE Ajout_LIAISON.Job=" & NmJob & ";"

    Con.Execute Sql
Sql = "DELETE Ajout_LIAISON_CONNECTEURS.*, Ajout_LIAISON_CONNECTEURS.Job FROM Ajout_LIAISON_CONNECTEURS "
    Sql = Sql & "WHERE Ajout_LIAISON_CONNECTEURS.Job=" & NmJob & ";"
    Con.Execute Sql
End Sub
