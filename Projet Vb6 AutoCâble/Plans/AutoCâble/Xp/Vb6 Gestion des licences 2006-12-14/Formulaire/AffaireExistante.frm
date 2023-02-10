VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AffaireExistante 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Affaire Existante"
   ClientHeight    =   4725
   ClientLeft      =   30
   ClientTop       =   300
   ClientWidth     =   9120
   Icon            =   "AffaireExistante.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OleObjectBlob   =   "AffaireExistante.dsx":030A
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "AffaireExistante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Annuler As Boolean
Private Sub CommandButton1_Click()
Dim Sql As String
Dim Rs As Recordset
CherchPices.Charge Me, "Pere=0", False
CherchPicesAnnuler = CherchPices.Annuler
Unload CherchPices
If CherchPicesAnnuler = True Then Exit Sub
If Trim("" & Me.txt3.Tag) = "" Then Exit Sub
Sql = "SELECT T_indiceProjet.Pere FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Me.txt3.Tag & ";"

Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = False Then
    If Rs!Pere > 0 Then Me.Tag = Rs!Pere
Else
    If Rs.EOF = True Then
    Sql = "SELECT Archive_T_indiceProjet.Pere FROM Archive_T_indiceProjet "
    Sql = Sql & "WHERE Archive_T_indiceProjet.Id=" & Me.txt3.Tag & ";"

    Set Rs = Con.OpenRecordSet(Sql)
        If Rs!Pere > 0 Then Me.Tag = Rs!Pere
    End If
End If
strStatus = ""
End Sub

Private Sub CommandButton2_Click()
Dim Piece As Long
Dim pathTmpXls As String
Dim Sql As String
Dim Rs As Recordset


Dim Fso As New FileSystemObject
Set FormBarGrah = Me
If Trim("" & Me.txt3.Tag) = "" Then
    MsgBox "Le champ Pièce est obligatoire"
    CommandButton1_Click
    Exit Sub
End If
Annuler = False
Me.Hide
End Sub

Private Sub CommandButton3_Click()
Me.Hide
End Sub

Private Sub UserForm_Activate()
Annuler = True
End Sub

