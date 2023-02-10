VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MajStockEboutique 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mis à jours du stock eboutique :"
   ClientHeight    =   5190
   ClientLeft      =   30
   ClientTop       =   435
   ClientWidth     =   9120
   Icon            =   "MajStockEboutiquel.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OleObjectBlob   =   "MajStockEboutiquel.dsx":08CA
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "MajStockEboutique"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public IdIndiceProjet As Long
Dim Id_Pere As Long
Dim Noquite As Boolean
Dim PreparNomOk As Integer
Dim NomenclatureOk As Boolean
Dim Id_Fils As Long
Dim Id_Projet As Long

Private Sub CommandButton1_Click()
If PreparNomOk <> 1 Then
    CherchPices.Charge Me, "(VerifieDate= Null   and Archiver=false) OR (IdStatus<4 and Archiver=false)"
Else
    CherchPices.Charge Me, "(VerifieDate= Null   and IdStatus<>4) "
End If
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
 
 If Trim("" & Me.txt3.Tag) = "" Then
    MsgBox "Le champ Pièce est obligatoire.", vbCritical, "Auto-Câble"
    CommandButton1_Click
    Exit Sub
End If

 If MyFormatQRY(Me.NbPieces) = False Then Exit Sub
 If MsgBox("Voulez-vous vraiment mètre à jour le Eboutique pour la pièce : " & Me.txt5, vbQuestion + vbYesNo, "Auto-Câble") = vbNo Then Exit Sub
 Id_Projet = Val(Me.txt3.Tag)
Sql = "SELECT T_indiceProjet.Pere FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Me.txt3.Tag & ";"
Set Rs = Con.OpenRecordSet(Sql)
If Rs!Pere > 0 Then
 Id_Fils = Me.txt3.Tag
   Id_Projet = Rs!Pere
Else
   Id_Fils = Id_Projet
    
End If
Set Rs = Con.CloseRecordSet(Rs)
If IsCilent = True Then
    Sql = "INSERT INTO T_Job ( [Action], Id_Piece, Id_Fils, Machine, Name, NbPieces )"
    Sql = Sql & "VALUES ('Maj Eboutique',  " & Id_Projet & ", " & Id_Fils & " , '" & MyReplace(UserName) & "' , '" & MyReplace(Me.txt5) & "' , " & Me.NbPieces & ");"
    Con.Execute Sql
    MsgBox "Votre demande a été prise en compte vous pouvez suivre l'évolution de votre travail dans la fenêtre (Liste des JOB)"
Else
    MajStock Id_Projet, Id_Fils, Val(NbPieces), Me
End If
 Noquite = False
 'frmAutocâble.DesEnabledMenu
Unload Me

End Sub

Private Sub CommandButton3_Click()
Noquite = False
'frmAutocâble.DesEnabledMenu
Unload Me
End Sub
Public Sub ChargeNomenclature(PreparNomk As Integer, MeCapTion As String)
Me.Caption = Me.Caption & " " & MeCapTion
    PreparNomOk = PreparNomk
    Me.Show vbModal
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

Private Sub UserForm_Activate()
 
Noquite = True
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Cancel = Noquite
End Sub
