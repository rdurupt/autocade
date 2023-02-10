VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Planche_Clous 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Critères de mise à jour :"
   ClientHeight    =   4935
   ClientLeft      =   30
   ClientTop       =   300
   ClientWidth     =   4455
   Icon            =   "Planche_Clous.dsx":0000
   MinButton       =   0   'False
   OleObjectBlob   =   "Planche_Clous.dsx":27A2
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Planche_Clous"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public boolAnnuler As Boolean
Dim boolCloseForm As Boolean
Dim MYPlanche As String
Dim Id As Long
Public Sub chargement(Id_Pieces As Long)
Dim Rs As Recordset
Dim Sql As String
'Set MyFrmApelant = FrmApelant
Id = Id_Pieces
Sql = "SELECT T_indiceProjet.Cartouche "
Sql = Sql & " FROM T_indiceProjet "
Sql = Sql & " WHERE T_indiceProjet.Id=" & Id_Pieces & ";"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = False Then
    aa = "" & Rs!Cartouche
    If Trim(aa) <> "" Then
    aa = Split(aa, "\")
    MYPlanche = aa(UBound(aa))

    For I = 0 To PlanchClous.ListCount - 1
    If UCase(PlanchClous.List(I)) = UCase(MYPlanche) Then PlanchClous.ListIndex = I
    Next
    End If
End If
Set Rs = Con.CloseRecordSet(Rs)
DoEvents
Me.Show vbModal
End Sub

Private Sub CommandButton1_Click()
Dim Sql As String
Dim Rs As Recordset
boolAnnuler = False
If Plan_Ouvrir.Value = True Or Outil_Ouvrir.Value = True Then
If Trim(PlanchClous.Text) = "" Then
    MsgBox "Vous devez sélectionner une planche à clous", vbExclamation
    Me.PlanchClous.SetFocus
    Exit Sub
End If
End If
 Sql = "UPDATE T_indiceProjet SET T_indiceProjet.Cartouche = '" & MyReplace(TableauPath.Item("PathOutils") & "\" & Me.PlanchClous) & "' "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Id & ";"
Con.Execute Sql
 bool_Plan_L_Connecteurs = Me.Plan_L_Connecteurs.Value
 bool_Plan_L_Fils = Me.Plan_L_Fils.Value
 bool_Plan_L_Vignettes = Me.Plan_L_Vignettes.Value
 bool_Plan_L_Etiquettes = Me.Plan_L_Etiquettes.Value
 bool_Plan_L_Composants = Me.Plan_L_Composants.Value
 bool_Plan_L_Notas = Me.Plan_L_Notas.Value
 bool_Plan_L_cartouches = Me.Plan_L_cartouches.Value
 bool_Plan_L_Preconisations = Me.Plan_L_Preconisations.Value
 bool_Plan_L_Options = Me.Plan_L_Options.Value
 bool_Plan_L_Criteres = Me.Plan_L_Criteres.Value
 bool_Plan_L_Noeuds = Me.Plan_L_Noeuds

 bool_Plan_E_Connecteurs = Me.Plan_E_Connecteurs.Value
 bool_Plan_E_Fils = Me.Plan_E_Fils.Value
 bool_Plan_E_Vignettes = Me.Plan_E_Vignettes.Value
 bool_Plan_E_Etiquettes = Me.Plan_E_Etiquettes.Value
 bool_Plan_E_Composants = Me.Plan_E_Composants.Value
 bool_Plan_E_Notas = Me.Plan_E_Notas.Value
 bool_Plan_E_cartouches = Me.Plan_E_cartouches.Value
 bool_Plan_E_Preconisations = Me.Plan_E_Preconisations.Value
 bool_Plan_E_Options = Me.Plan_E_Options.Value
 bool_Plan_E_Criteres = Me.Plan_E_Criteres.Value
 bool_Plan_E_Noeuds = Me.Plan_E_Noeuds


 bool_Outil_L_Connecteurs = Me.Outil_L_Connecteurs.Value
 bool_Outil_L_Fils = Me.Outil_L_Fils.Value
 bool_Outil_L_Vignettes = Me.Outil_L_Vignettes.Value
 bool_Outil_L_Etiquettes = Me.Outil_L_Etiquettes.Value
 bool_Outil_L_Composants = Me.Outil_L_Composants.Value
 bool_Outil_L_Notas = Me.Outil_L_Notas.Value
 bool_Outil_L_cartouches = Me.Outil_L_cartouches.Value
 bool_Outil_L_Preconisations = Outil_L_Preconisations.Value
 bool_Outil_L_Options = Me.Outil_L_Options.Value
 bool_Outil_L_Criteres = Me.Outil_L_Criteres.Value
 bool_Outil_L_Noeuds = Me.Outil_L_Noeuds
 
 
 bool_Outil_E_Connecteurs = Me.Outil_E_Connecteurs.Value
 bool_Outil_E_Fils = Me.Outil_E_Fils.Value
 bool_Outil_E_Vignettes = Me.Outil_E_Vignettes.Value
 bool_Outil_E_Etiquettes = Me.Outil_E_Etiquettes.Value
 bool_Outil_E_Composants = Me.Outil_E_Composants.Value
 bool_Outil_E_Notas = Me.Outil_E_Notas.Value
 bool_Outil_E_cartouches = Me.Outil_E_cartouches.Value
 bool_Outil_E_Preconisations = Me.Outil_E_Preconisations.Value
 bool_Outil_E_Options = Me.Outil_E_Options.Value
 bool_Outil_E_Criteres = Me.Outil_E_Criteres.Value
 bool_Outil_E_Noeuds = Me.Outil_E_Noeuds

 bool_Plan_Ouvrir = Me.Plan_Ouvrir.Value
 bool_Outil_Ouvrir = Me.Outil_Ouvrir.Value


If (bool_Plan_Ouvrir Or bool_Outil_Ouvrir) = False Then boolAnnuler = True
If Plan_Ouvrir.Value = True Or Outil_Ouvrir.Value = True Then
If IsCilent = True Then
Sql = "SELECT [PI] & '_' & Trim([PI_Indice]) AS Name  "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Id & ";"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = False Then

    Sql = "DELETE T_Job.* FROM T_Job "
    Sql = Sql & "WHERE T_Job.Id_Piece=" & Id & ";"
    Con.Execute Sql
    
    Sql = "INSERT INTO T_Job ( Id_Piece, Id_Fils, Action,Outil_E_cartouches, Outil_E_Connecteurs, Outil_E_Criteres,  "
    Sql = Sql & "Outil_E_Etiquettes, Outil_E_Fils, Outil_E_Noeuds, Outil_E_Notas, Outil_E_Options,  "
    Sql = Sql & "Outil_E_Preconisations, Outil_E_Vignettes, Outil_L_cartouches, Outil_L_Composants,  "
    Sql = Sql & "Outil_L_Connecteurs, Outil_L_Criteres, Outil_L_Etiquettes, Outil_L_Fils, Outil_L_Noeuds,  "
    Sql = Sql & "Outil_L_Notas, Outil_L_Options, Outil_L_Preconisations, Outil_L_Vignettes, Outil_Ouvrir,  "
    Sql = Sql & "Plan_E_cartouches, Plan_E_Composants, Plan_E_Connecteurs, Plan_E_Criteres, Plan_E_Etiquettes,  "
    Sql = Sql & "Plan_E_Fils, Plan_E_Noeuds, Plan_E_Notas, Plan_E_Options, Plan_E_Preconisations, Plan_E_Vignettes,  "
    Sql = Sql & "Plan_L_cartouches, Plan_L_Composants, Plan_L_Connecteurs, Plan_L_Criteres, Plan_L_Etiquettes,  "
    Sql = Sql & "Plan_L_Fils, Plan_L_Noeuds, Plan_L_Notas, Plan_L_Options, Plan_L_Preconisations, Plan_L_Vignettes,  "
    Sql = Sql & "Plan_Ouvrir,Outil_E_Composants, Machine,Name )"
     Sql = Sql & "VALUES (" & Id & ", " & IdFils & ",'Modifier Plan', " & MyReplaceBool(Me.Outil_E_cartouches) & ", " & MyReplaceBool(Me.Outil_E_Connecteurs) & ", " & MyReplaceBool(Me.Outil_E_Criteres) & ",  "
    Sql = Sql & "" & MyReplaceBool(Me.Outil_E_Etiquettes) & ", " & MyReplaceBool(Me.Outil_E_Fils) & ", " & MyReplaceBool(Me.Outil_E_Noeuds) & ", " & MyReplaceBool(Me.Outil_E_Notas) & ",  "
    Sql = Sql & "" & MyReplaceBool(Me.Outil_E_Options) & ", " & MyReplaceBool(Me.Outil_E_Preconisations) & ", " & MyReplaceBool(Me.Outil_E_Vignettes) & ", " & MyReplaceBool(Me.Outil_L_cartouches) & ",  "
    Sql = Sql & "" & MyReplaceBool(Me.Outil_L_Composants) & ", " & MyReplaceBool(Me.Outil_L_Connecteurs) & ", " & MyReplaceBool(Me.Outil_L_Criteres) & ", " & MyReplaceBool(Me.Outil_L_Etiquettes) & ",  "
    Sql = Sql & "" & MyReplaceBool(Me.Outil_L_Fils) & ", " & MyReplaceBool(Me.Outil_L_Noeuds) & ", " & MyReplaceBool(Me.Outil_L_Notas) & ", " & MyReplaceBool(Me.Outil_L_Options) & ", " & MyReplaceBool(Me.Outil_L_Preconisations) & ",  "
    Sql = Sql & "" & MyReplaceBool(Me.Outil_L_Vignettes) & ", " & MyReplaceBool(Me.Outil_Ouvrir) & ", " & MyReplaceBool(Me.Plan_E_cartouches) & ", " & MyReplaceBool(Me.Plan_E_Composants) & ", " & MyReplaceBool(Me.Plan_E_Connecteurs) & ",  "
    Sql = Sql & "" & MyReplaceBool(Me.Plan_E_Criteres) & ", " & MyReplaceBool(Me.Plan_E_Etiquettes) & ", " & MyReplaceBool(Me.Plan_E_Fils) & ", " & MyReplaceBool(Me.Plan_E_Noeuds) & ", " & MyReplaceBool(Me.Plan_E_Notas) & ", " & MyReplaceBool(Me.Plan_E_Options) & ",  "
    Sql = Sql & "" & MyReplaceBool(Me.Plan_E_Preconisations) & ", " & MyReplaceBool(Me.Plan_E_Vignettes) & ", " & MyReplaceBool(Me.Plan_L_cartouches) & ", " & MyReplaceBool(Me.Plan_L_Composants) & ",  "
    Sql = Sql & "" & MyReplaceBool(Me.Plan_L_Connecteurs) & ", " & MyReplaceBool(Me.Plan_L_Criteres) & ", " & MyReplaceBool(Me.Plan_L_Etiquettes) & ", " & MyReplaceBool(Me.Plan_L_Fils) & ",  "
    Sql = Sql & "" & MyReplaceBool(Me.Plan_L_Noeuds) & ", " & MyReplaceBool(Me.Plan_L_Notas) & ", " & MyReplaceBool(Me.Plan_L_Options) & ", " & MyReplaceBool(Me.Plan_L_Preconisations) & ", " & MyReplaceBool(Me.Plan_L_Vignettes) & ", " & MyReplaceBool(Me.Plan_Ouvrir) & "," & MyReplaceBool(Me.Outil_E_Composants) & ", '" & MyReplace(UserName) & "','" & MyReplace("" & Rs!Name) & "' );"
Set Rs = Con.CloseRecordSet(Rs)
End If





Con.Execute Sql
End If
End If
boolCloseForm = False
Me.Hide
End Sub

Private Sub CommandButton2_Click()
boolAnnuler = True
boolCloseForm = False
Me.Hide
End Sub

Private Sub Outil_L_cartouches_Click()
Outil_E_cartouches.Value = Outil_L_cartouches.Value

End Sub

Private Sub Outil_L_Composants_Click()
Outil_E_Composants.Value = Outil_L_Composants.Value

End Sub

Private Sub Outil_L_Fils_Click()
Outil_E_Fils.Value = Outil_L_Fils.Value
Outil_L_Etiquettes.Value = Outil_L_Fils.Value
Outil_E_Etiquettes.Value = Outil_L_Fils.Value
Outil_L_Preconisations.Value = Outil_L_Fils.Value
Outil_E_Preconisations.Value = Outil_L_Fils.Value
Outil_L_Options.Value = Outil_L_Fils.Value
Outil_L_Criteres.Value = Outil_L_Fils.Value
Outil_E_Criteres.Value = Outil_L_Fils.Value
Outil_E_Connecteurs.Value = Outil_L_Fils.Value
Outil_L_Vignettes.Value = Outil_L_Fils.Value
Outil_E_Vignettes.Value = Outil_L_Fils.Value
Outil_L_Connecteurs.Value = Outil_L_Fils.Value
End Sub

Private Sub Outil_L_Noeuds_Click()
Outil_E_Noeuds.Value = Outil_L_Noeuds.Value

End Sub

Private Sub Outil_L_Notas_Click()
 Outil_E_Notas.Value = Outil_L_Notas.Value
End Sub

Private Sub Plan_L_cartouches_Click()
Plan_E_cartouches.Value = Plan_L_cartouches.Value
End Sub

Private Sub Plan_L_Composants_Click()
Plan_E_Composants.Value = Plan_L_Composants.Value
End Sub

Private Sub Plan_L_Fils_Click()
Plan_E_Fils.Value = Plan_L_Fils.Value
Plan_L_Etiquettes.Value = Plan_L_Fils.Value
Plan_E_Etiquettes.Value = Plan_L_Fils.Value
Plan_L_Preconisations.Value = Plan_L_Fils.Value
Plan_E_Preconisations.Value = Plan_L_Fils.Value
Plan_L_Options.Value = Plan_L_Fils.Value
Plan_L_Criteres.Value = Plan_L_Fils.Value
Plan_E_Criteres.Value = Plan_L_Fils.Value
Plan_E_Connecteurs.Value = Plan_L_Fils.Value
Plan_L_Vignettes.Value = Plan_L_Fils.Value
Plan_E_Vignettes.Value = Plan_L_Fils.Value
Plan_L_Connecteurs.Value = Plan_L_Fils.Value
End Sub

Private Sub Plan_L_Noeuds_Click()
Plan_E_Noeuds.Value = Plan_L_Noeuds.Value
End Sub

Private Sub Plan_L_Notas_Click()
Plan_E_Notas.Value = Plan_L_Notas.Value
End Sub

Private Sub UserForm_Initialize()
Dim Sql As String
Dim MyPath As String
Dim Rs As Recordset
Dim MyFichier As String
Set TableauPath = funPath
PlanchClous.Clear
MyPath = TableauPath.Item("PathOutils") & "\"
MyPath = DefinirChemienComplet(TableauPath.Item("PathServer"), MyPath)

'If Left(MyPath, 2) <> "\\" And Left(MyPath, 1) = "\" Then MyPath = TableauPath.Item("PathServer") & MyPath & "\"
'If Right(MyPath, 2) = "\\" Then MyPath = Mid(MyPath, 1, Len(MyPath) - 1)



If Trim(MyPath) <> "" Then
MyFichier = Dir(MyPath & "*.dwg")
PlanchClous.AddItem ""
While MyFichier <> ""
PlanchClous.AddItem MyFichier
    MyFichier = Dir
 Wend
End If
boolCloseForm = True
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Cancel = boolCloseForm
End Sub
