VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExporterExcel 
   Caption         =   "Exporter Excel :"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9120
   Icon            =   "ExporterExcel.dsx":0000
   OleObjectBlob   =   "ExporterExcel.dsx":08CA
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ExporterExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public IdIndiceProjet As Long
Dim Id_Pere As Long
Dim Noquite As Boolean

Private Sub CommandButton1_Click()
CherchPices.Charge Me, "(VerifieDate= Null  and Archiver=False) OR (IdStatus=3 and Archiver=False)"
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
    
Sql = "SELECT T_indiceProjet.Pere FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Me.txt3.Tag & ";"
Set Rs = Con.OpenRecordSet(Sql)
If Rs!Pere > 0 Then
IdFils = Me.txt3.Tag
    Me.txt3.Tag = Rs!Pere
Else
   IdFils = 0
    
End If
Set FormBarGrah = Me



subExporteXls Me.txt3.Tag
 Noquite = False
Me.Hide

End Sub

Private Sub CommandButton3_Click()
Noquite = False
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
For i = 0 To 11
    Me.Controls("txt" & CStr(i + 1)) = "" & Rs(i)
     Me.Controls("txt" & CStr(i + 1)).Tag = "" & Rs.Fields(12)

Next i
    
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
