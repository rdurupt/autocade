VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmEtiquette 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Générateur d'étiquettes :"
   ClientHeight    =   6225
   ClientLeft      =   30
   ClientTop       =   300
   ClientWidth     =   9225
   Icon            =   "FrmEtiquettel.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OleObjectBlob   =   "FrmEtiquettel.dsx":030A
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "FrmEtiquette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public IdIndiceProjet As Long
Dim Id_Pere As Long
Dim Noquite As Boolean

Private Sub CommandButton1_Click()
CherchPices.Charge Me, "(VerifieDate= Null    and Archiver=false) OR (IdStatus<4 and Archiver=false)"
Unload CherchPices
End Sub

Private Sub CommandButton2_Click()
Dim Piece As Long
Dim pathTmpXls As String
Dim Sql As String
Dim Rs As Recordset
Dim Fso As New FileSystemObject
Dim UserForm2_boolExcute As Boolean
Dim Planche_Clous_boolAnnuler As Boolean
Dim Equipement
Dim Equipement2
Dim Equipement3 As String
If Trim("" & Me.txt3.Tag) = "" Then
    MsgBox "Le champ Pièce est obligatoire.", vbCritical, "Auto-Câble"
    CommandButton1_Click
    Exit Sub
End If
    
Sql = "SELECT T_indiceProjet.Pere, T_indiceProjet.Equipement FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Me.txt3.Tag & ";"
Set Rs = Con.OpenRecordSet(Sql)
Equipement = "" & Rs!Equipement
If Rs!Pere > 0 Then
IdFils = Me.txt3.Tag
    Me.txt3.Tag = Rs!Pere
Else
   IdFils = 0
    
End If
Set FormBarGrah = Me
Equipement = Split(Equipement & ";", ";")
For I = 0 To UBound(Equipement)
    Equipement2 = Split(Equipement(I) & "_", "_")
    If Trim("" & Equipement2(0)) <> "" Then
        Equipement3 = Equipement3 & ";" & Equipement2(0) & ";"
    End If
Next
Me.Enabled = False


If IsCilent = True Then
Sql = "SELECT [PI] & '_' & Trim([PI_Indice]) AS Name  "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Me.txt3.Tag & ";"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = False Then

    Sql = "DELETE T_Job.* FROM T_Job "
    Sql = Sql & "WHERE T_Job.Id_Piece=" & Me.txt3.Tag & ";"
    Con.Execute Sql
    
Set Rs = Con.CloseRecordSet(Rs)
Sql = "INSERT INTO T_Job ( [Action], Id_Piece, Id_Fils, Machine, Name, Nomenclature_Appareil, Par_Fournisseur, Par_Options )"
Sql = Sql & "VALUES('Créer Ettiquettes' , " & Me.txt3.Tag & " , " & IdFils & " , '" & MyReplace(UserName) & "' , '" & MyReplace(txt5) & "' , " & OptionButton4.Value * 1 & " , " & CheckBox2.Value * 1 & " , " & CheckBox1.Value * 1 & " );"
Con.Execute Sql

MsgBox "Votre demande a été prise en compte vous pouvez suivre l'évolution de votre travail dans la fenêtre (Liste des JOB)"
End If






Else





If OptionButton4.Value = True Then
  GenairEtiquette2 Val(Me.txt3.Tag), Equipement3, CheckBox1.Value, CheckBox2.Value
Else
    GenairEtiquette Val(Me.txt3.Tag)
End If
End If
 Noquite = False
 'frmAutocâble.DesEnabledMenu
Me.Hide

End Sub

Private Sub CommandButton3_Click()
Noquite = False
'frmAutocâble.DesEnabledMenu
Me.Hide
End Sub
Public Sub Charge(MyForm As Object)
Dim Sql As String
Dim Rs As Recordset
IdIndiceProjet = MyForm.IdIndiceProjet

Sql = "SELECT SelectProjets.* FROM SelectProjets WHERE SelectProjets.Id=" & IdIndiceProjet & " ;"

Set Rs = Con.OpenRecordSet(Sql)

Set FormBarGrah = Me
If Rs.EOF = False Then
For I = 0 To 11
    Me.Controls("txt" & CStr(I + 1)) = "" & Rs(I)
     Me.Controls("txt" & CStr(I + 1)).Tag = "" & Rs.Fields(12)

Next I
    
    OptionButton2.Value = True
    OptionButton1.Value = False
    Me.CommandButton1.Enabled = True
 End If
 Set Rs = Con.CloseRecordSet(Rs)
 MyForm.Hide
 Me.Show vbModal
End Sub

Private Sub OptionButton1_Click()
OptionButton2.Value = False
End Sub

Private Sub OptionButton2_Click()
OptionButton1.Value = False
End Sub

Private Sub OptionButton3_Click()
If OptionButton3.Value = True Then
    OptionButton4.Value = 0
    CheckBox1.Enabled = False
    CheckBox2.Enabled = False
End If
End Sub

Private Sub OptionButton4_Click()
If OptionButton4.Value = True Then
    OptionButton3.Value = 0
    CheckBox1.Enabled = True
    CheckBox2.Enabled = True
End If
End Sub

Private Sub UserForm_Activate()
Noquite = True
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Cancel = Noquite
End Sub
